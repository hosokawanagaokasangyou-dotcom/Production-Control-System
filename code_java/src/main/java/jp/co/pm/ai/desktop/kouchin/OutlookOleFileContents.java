package jp.co.pm.ai.desktop.kouchin;

import java.io.ByteArrayOutputStream;
import java.util.List;
import java.util.Locale;

import com.sun.jna.Memory;
import com.sun.jna.Native;
import com.sun.jna.Pointer;
import com.sun.jna.Structure;
import com.sun.jna.platform.win32.Ole32;
import com.sun.jna.platform.win32.WinNT.HRESULT;
import com.sun.jna.ptr.IntByReference;
import com.sun.jna.ptr.PointerByReference;
import com.sun.jna.win32.StdCallLibrary;
import com.sun.jna.win32.W32APIOptions;

/**
 * Outlook 添付の本体。JavaFX の Dragboard.getContent は null になるため、
 * OLE の FileContents（IStream / HGLOBAL）を読む。
 */
final class OutlookOleFileContents {

    private static final int DVASPECT_CONTENT = 1;
    private static final int TYMED_HGLOBAL = 1;
    private static final int TYMED_ISTREAM = 4;
    private static final int S_OK = 0;
    private static final int MAX_BYTES = 32 * 1024 * 1024;

    record Result(byte[] bytes, String detail) {}

    private OutlookOleFileContents() {}

    static Result read() {
        String os = System.getProperty("os.name", "");
        if (!os.toLowerCase(Locale.ROOT).contains("win")) {
            return new Result(null, "not-windows");
        }
        StringBuilder detail = new StringBuilder();
        byte[] awt = readAwtFileList(detail);
        if (awt != null && awt.length > 0) {
            return new Result(awt, detail.toString());
        }
        HRESULT init = Ole32.INSTANCE.OleInitialize(null);
        boolean owned = init != null && init.intValue() == S_OK;
        detail.append("init=").append(hex(init == null ? -1 : init.intValue())).append(' ');
        try {
            int format = NativeFormats.INSTANCE.RegisterClipboardFormat("FileContents");
            detail.append("cf=").append(format);
            if (format == 0) {
                return new Result(null, detail.toString());
            }
            PointerByReference ref = new PointerByReference();
            HRESULT clip = NativeOle.INSTANCE.OleGetClipboard(ref);
            int clipHr = clip == null ? -1 : clip.intValue();
            detail.append(" clip=").append(hex(clipHr));
            Pointer data = ref.getValue();
            if (clipHr != S_OK || data == null) {
                return new Result(null, detail.toString());
            }
            try {
                for (int lindex : new int[] {0, -1}) {
                    for (int tymed : new int[] {TYMED_ISTREAM, TYMED_HGLOBAL}) {
                        byte[] got = getData(data, format, lindex, tymed, detail);
                        if (got != null && got.length > 0) {
                            return new Result(got, detail.toString());
                        }
                    }
                }
            } finally {
                releaseUnknown(data);
            }
            return new Result(null, detail.toString());
        } catch (Throwable ex) {
            detail.append(" ex=").append(ex.getClass().getSimpleName());
            return new Result(null, detail.toString());
        } finally {
            if (owned) {
                Ole32.INSTANCE.OleUninitialize();
            }
        }
    }

    private static byte[] readAwtFileList(StringBuilder detail) {
        try {
            java.awt.datatransfer.Clipboard awt = java.awt.Toolkit.getDefaultToolkit().getSystemClipboard();
            java.awt.datatransfer.DataFlavor[] flavors = awt.getAvailableDataFlavors();
            detail.append("awt=").append(flavors == null ? -1 : flavors.length);
            if (flavors == null) {
                return null;
            }
            if (!awt.isDataFlavorAvailable(java.awt.datatransfer.DataFlavor.javaFileListFlavor)) {
                return null;
            }
            Object data = awt.getData(java.awt.datatransfer.DataFlavor.javaFileListFlavor);
            if (!(data instanceof List<?> list) || list.isEmpty() || !(list.get(0) instanceof java.io.File file)) {
                return null;
            }
            if (!file.isFile() || file.length() <= 0 || file.length() > MAX_BYTES) {
                return null;
            }
            detail.append(" awtFile=").append(file.length());
            return java.nio.file.Files.readAllBytes(file.toPath());
        } catch (Throwable ex) {
            detail.append(" awtEx=").append(ex.getClass().getSimpleName());
            return null;
        }
    }

    private static byte[] getData(Pointer data, int format, int lindex, int tymed, StringBuilder detail) {
        FormatEtc etc = new FormatEtc();
        etc.cfFormat = (short) format;
        etc.dwAspect = DVASPECT_CONTENT;
        etc.lindex = lindex;
        etc.tymed = tymed;
        etc.write();
        StgMedium medium = new StgMedium();
        medium.write();
        int hr = invokeGetData(data, etc, medium);
        medium.read();
        detail.append(" l=").append(lindex).append(" t=").append(tymed).append(" hr=").append(hex(hr));
        if (hr != S_OK) {
            return null;
        }
        try {
            if (medium.tymed == TYMED_ISTREAM && medium.pData != null) {
                byte[] bytes = readStream(medium.pData);
                detail.append(" stream=").append(bytes == null ? -1 : bytes.length);
                return bytes;
            }
            if (medium.tymed == TYMED_HGLOBAL && medium.pData != null) {
                byte[] bytes = readGlobal(medium.pData);
                detail.append(" global=").append(bytes == null ? -1 : bytes.length);
                return bytes;
            }
            detail.append(" tymed=").append(medium.tymed);
            return null;
        } finally {
            try {
                NativeOle.INSTANCE.ReleaseStgMedium(medium);
            } catch (Throwable ignored) {
            }
        }
    }

    private static int invokeGetData(Pointer data, FormatEtc etc, StgMedium medium) {
        Pointer vtbl = data.getPointer(0);
        Pointer fn = vtbl.getPointer(3L * Native.POINTER_SIZE);
        com.sun.jna.Function getData = com.sun.jna.Function.getFunction(fn, com.sun.jna.Function.ALT_CONVENTION);
        return getData.invokeInt(new Object[] {data, etc, medium});
    }

    private static byte[] readStream(Pointer stream) {
        Pointer vtbl = stream.getPointer(0);
        com.sun.jna.Function read = com.sun.jna.Function.getFunction(
                vtbl.getPointer(3L * Native.POINTER_SIZE), com.sun.jna.Function.ALT_CONVENTION);
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        Memory mem = new Memory(64 * 1024);
        IntByReference got = new IntByReference();
        while (out.size() < MAX_BYTES) {
            got.setValue(0);
            int hr = read.invokeInt(new Object[] {stream, mem, (int) mem.size(), got});
            int n = got.getValue();
            if (hr != S_OK || n <= 0) {
                break;
            }
            int take = Math.min(n, MAX_BYTES - out.size());
            out.write(mem.getByteArray(0, take), 0, take);
            if (n < mem.size()) {
                break;
            }
        }
        return out.size() == 0 ? null : out.toByteArray();
    }

    private static byte[] readGlobal(Pointer global) {
        int n = NativeKernel.INSTANCE.GlobalSize(global);
        if (n <= 0 || n > MAX_BYTES) {
            return null;
        }
        Pointer locked = NativeKernel.INSTANCE.GlobalLock(global);
        if (locked == null) {
            return null;
        }
        try {
            return locked.getByteArray(0, n);
        } finally {
            NativeKernel.INSTANCE.GlobalUnlock(global);
        }
    }

    private static void releaseUnknown(Pointer unknown) {
        Pointer vtbl = unknown.getPointer(0);
        com.sun.jna.Function release = com.sun.jna.Function.getFunction(
                vtbl.getPointer(2L * Native.POINTER_SIZE), com.sun.jna.Function.ALT_CONVENTION);
        release.invokeInt(new Object[] {unknown});
    }

    private static String hex(int hr) {
        return Integer.toHexString(hr);
    }

    private interface NativeKernel extends StdCallLibrary {
        NativeKernel INSTANCE = Native.load("kernel32", NativeKernel.class, W32APIOptions.DEFAULT_OPTIONS);

        Pointer GlobalLock(Pointer hMem);

        boolean GlobalUnlock(Pointer hMem);

        int GlobalSize(Pointer hMem);
    }

    private interface NativeOle extends StdCallLibrary {
        NativeOle INSTANCE = Native.load("ole32", NativeOle.class, W32APIOptions.DEFAULT_OPTIONS);

        HRESULT OleGetClipboard(PointerByReference ppDataObj);

        void ReleaseStgMedium(StgMedium medium);
    }

    private interface NativeFormats extends StdCallLibrary {
        NativeFormats INSTANCE = Native.load("user32", NativeFormats.class, W32APIOptions.DEFAULT_OPTIONS);

        int RegisterClipboardFormat(String lpszFormat);
    }

    public static class FormatEtc extends Structure {
        public short cfFormat;
        public Pointer ptd;
        public int dwAspect;
        public int lindex;
        public int tymed;

        @Override
        protected List<String> getFieldOrder() {
            return List.of("cfFormat", "ptd", "dwAspect", "lindex", "tymed");
        }
    }

    public static class StgMedium extends Structure {
        public int tymed;
        public Pointer pData;
        public Pointer pUnkForRelease;

        @Override
        protected List<String> getFieldOrder() {
            return List.of("tymed", "pData", "pUnkForRelease");
        }
    }
}
