package jp.co.pm.ai.desktop.ui;

import javafx.scene.control.Slider;

/**
 * スライダーのドラッグ中はラベル更新のみにし、確定時（ドラッグ終了・トラッククリック・キー調整など {@link
 * Slider#isValueChanging()} が false の変化）に重い処理を走らせる。
 */
public final class SliderCommittedChangeSupport {

    private SliderCommittedChangeSupport() {}

    /**
     * @param updateLabel 値に応じたラベル表示など（ドラッグ中も毎回）。{@code null} 可。
     * @param onCommitted 永続化・再描画など。ドラッグ終了時およびトラッククリック時に実行。
     */
    public static void install(Slider slider, Runnable updateLabel, Runnable onCommitted) {
        if (slider == null || onCommitted == null) {
            return;
        }
        Runnable label = updateLabel != null ? updateLabel : () -> {};
        slider.valueProperty()
                .addListener(
                        (o, a, b) -> {
                            label.run();
                            if (!slider.isValueChanging()) {
                                onCommitted.run();
                            }
                        });
        slider.valueChangingProperty()
                .addListener(
                        (obs, wasChanging, changing) -> {
                            if (Boolean.FALSE.equals(changing)) {
                                onCommitted.run();
                            }
                        });
    }
}
