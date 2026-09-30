package jp.co.pm.ai.desktop.kouchin;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Locale;

import com.sun.jna.platform.win32.COM.util.Factory;
import com.sun.jna.platform.win32.COM.util.IDispatch;
import com.sun.jna.platform.win32.COM.util.annotation.ComObject;

/**
 * JavaFX も OLE クリップボードも添付本体を返さない。
 * ドラッグで分かったファイル名を、起動中の Outlook の選択メールから保存する。
 */
public final class OutlookRunningAttachment {

    record Result(Path path, String detail) {}

    @ComObject(progId = "Outlook.Application", clsId = "0006F03A-0000-0000-C000-000000000046")
    interface Application extends IDispatch {}

    private OutlookRunningAttachment() {}

    static Result saveMatch(List<String> wantedNames, Path tempDir) {
        String os = System.getProperty("os.name", "");
        if (!os.toLowerCase(Locale.ROOT).contains("win")) {
            return new Result(null, "not-windows");
        }
        String wanted = firstName(wantedNames);
        if (wanted.isEmpty() || tempDir == null) {
            return new Result(null, "no-name");
        }
        Factory factory = new Factory();
        StringBuilder detail = new StringBuilder("name=").append(wanted);
        try {
            Application app = factory.fetchObject(Application.class);
            if (app == null) {
                return new Result(null, detail.append(" app=null").toString());
            }
            IDispatch item = selectedItem(app, detail);
            if (item == null) {
                return new Result(null, detail.append(" item=null").toString());
            }
            IDispatch attachments = item.getProperty(IDispatch.class, "Attachments");
            int count = attachments == null ? 0 : asInt(attachments.getProperty(Object.class, "Count"));
            detail.append(" att=").append(count);
            for (int i = 1; i <= count && i <= 32; i++) {
                IDispatch att = attachments.invokeMethod(IDispatch.class, "Item", i);
                if (att == null) {
                    continue;
                }
                String fileName = att.getProperty(String.class, "FileName");
                if (fileName == null || !fileName.equalsIgnoreCase(wanted)) {
                    continue;
                }
                Path dir = tempDir.resolve("pm-ai-outlook-drop");
                Files.createDirectories(dir);
                String safe = KouchinOutlookDropSupport.sanitizeDropFileName(fileName);
                Path dest = dir.resolve(safe);
                att.invokeMethod(Void.class, "SaveAsFile", dest.toAbsolutePath().toString());
                long size = Files.isRegularFile(dest) ? Files.size(dest) : -1;
                detail.append(" saved=").append(size);
                if (size <= 0) {
                    return new Result(null, detail.toString());
                }
                return new Result(dest.toAbsolutePath().normalize(), detail.toString());
            }
            return new Result(null, detail.append(" match=0").toString());
        } catch (Throwable ex) {
            return new Result(null, detail.append(" ex=").append(ex.getClass().getSimpleName()).toString());
        } finally {
            try {
                factory.disposeAll();
            } catch (Throwable ignored) {
            }
        }
    }

    private static IDispatch selectedItem(Application app, StringBuilder detail) {
        IDispatch explorer = app.invokeMethod(IDispatch.class, "ActiveExplorer");
        if (explorer != null) {
            IDispatch selection = explorer.getProperty(IDispatch.class, "Selection");
            int n = selection == null ? 0 : asInt(selection.getProperty(Object.class, "Count"));
            detail.append(" sel=").append(n);
            if (n > 0) {
                IDispatch item = selection.invokeMethod(IDispatch.class, "Item", 1);
                if (item != null) {
                    return item;
                }
            }
        } else {
            detail.append(" sel=0");
        }
        IDispatch inspector = app.getProperty(IDispatch.class, "ActiveInspector");
        if (inspector == null) {
            return null;
        }
        detail.append(" inspector=1");
        return inspector.getProperty(IDispatch.class, "CurrentItem");
    }

    private static String firstName(List<String> names) {
        if (names == null) {
            return "";
        }
        for (String name : names) {
            if (name != null && !name.isBlank()) {
                return PathName.fileName(name);
            }
        }
        return "";
    }

    private static int asInt(Object value) {
        if (value instanceof Number number) {
            return number.intValue();
        }
        return 0;
    }

    private static final class PathName {
        private static String fileName(String raw) {
            String name = raw.trim();
            int slash = Math.max(name.lastIndexOf('/'), name.lastIndexOf('\\'));
            if (slash >= 0 && slash + 1 < name.length()) {
                name = name.substring(slash + 1);
            }
            return name;
        }
    }
}
