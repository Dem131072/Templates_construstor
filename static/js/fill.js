/* Страница заполнения DOCX-шаблона.
 *
 * Удобный ввод значений с отображением изменений, которые будут применены, в предпросмотре.
 */

document.addEventListener("DOMContentLoaded", () => {
    const root = document.getElementById("fill-page");
    const filename = root?.dataset.filename || "";
    const form = document.getElementById("fill-form");
    const emptyFieldsModalElement = document.getElementById("emptyFieldsModal");
    const confirmSaveEmptyFieldsBtn = document.getElementById("confirm-save-empty-fields");
    const emptyFieldsModal = emptyFieldsModalElement ? new bootstrap.Modal(emptyFieldsModalElement) : null;
    let allowSubmitWithEmptyFields = false;
    const escapeHtml = (value) =>
        String(value ?? "")
            .replaceAll("&", "&amp;")
            .replaceAll("<", "&lt;")
            .replaceAll(">", "&gt;")
            .replaceAll('"', "&quot;")
            .replaceAll("'", "&#39;");
    const normalizeNumber = (value) => String(value ?? "").trim().replace(",", ".");
    const parseFloatLoose = (value) => {
        const normalized = normalizeNumber(value);
        if (!normalized) return null;
        const parsed = Number(normalized);
        return Number.isFinite(parsed) ? parsed : null;
    };
    const formatFloatTrimmed = (value) => {
        if (!Number.isFinite(value)) return "";
        let s = value.toFixed(10);
        s = s.replace(/0+$/, "").replace(/\.$/, "");
        return s;
    };
    const HARDWARE_ERROR_FACTOR = 4.0;
    const HARDWARE_ERROR_DIVISOR = 3.4641;
    const DEFAULT_MAIN_RELATIVE_ERROR_PERCENT = "8";
    const DEFAULT_ADDITIONAL_RELATIVE_ERROR_PERCENT = "0";

    // Формула аппаратной погрешности при измерении освещенности
    // Нужна только для отображения на шаблоне
    const calculateHardwareError = (
        measurementResult,
        mainRelativeErrorPercent,
        additionalRelativeErrorPercent
    ) => HARDWARE_ERROR_FACTOR * Math.sqrt(
        Math.pow(((mainRelativeErrorPercent / 100.0 * measurementResult) / HARDWARE_ERROR_DIVISOR), 2) +
        Math.pow(((additionalRelativeErrorPercent / 100.0 * measurementResult) / HARDWARE_ERROR_DIVISOR), 2)
    );
    const formatHardwareErrorValue = (
        measurementResult,
        mainRelativeErrorPercent,
        additionalRelativeErrorPercent
    ) => {
        const error = calculateHardwareError(
            measurementResult,
            mainRelativeErrorPercent,
            additionalRelativeErrorPercent
        );
        return `${formatFloatTrimmed(measurementResult)}±${error.toFixed(2)}`;
    };

    // Обновляет все спаны одного плейсхолдера в предпросмотре
    const setPreviewValue = (ph, fallbackName, renderedValue) => {
        const placeholders = document.querySelectorAll(`.placeholder[data-ph="${CSS.escape(ph)}"]`);
        placeholders.forEach(span => {
            const hasValue = String(renderedValue ?? "").trim() !== "";
            span.textContent = hasValue ? renderedValue : fallbackName;
            span.classList.toggle("placeholder-filled", hasValue);
        });
    };
    const syncRegularFieldPreview = (field) => {
        const ph = field.dataset.ph;
        const fallbackName = field.dataset.placeholderName || ph;
        const control = field.querySelector('input:not(.hardware-error-input), select, textarea');
        if (!control) return;
        setPreviewValue(ph, fallbackName, control.value.trim());
    };
    const syncHardwareErrorPreview = (field) => {
        const ph = field.dataset.ph;
        const fallbackName = field.dataset.placeholderName || ph;
        const measurementInput = field.querySelector('[data-role="measurement_result"]');
        const mainInput = field.querySelector('[data-role="main_relative_error_percent"]');
        const additionalInput = field.querySelector('[data-role="additional_relative_error_percent"]');
        const measurement = parseFloatLoose(measurementInput?.value ?? "");
        const main = parseFloatLoose(mainInput?.value ?? DEFAULT_MAIN_RELATIVE_ERROR_PERCENT);
        const additional = parseFloatLoose(additionalInput?.value ?? DEFAULT_ADDITIONAL_RELATIVE_ERROR_PERCENT);
        if (measurement === null || main === null || additional === null) {
            setPreviewValue(ph, fallbackName, "");
            return;
        }
        const result = formatHardwareErrorValue(measurement, main, additional);
        setPreviewValue(ph, fallbackName, result);
    };
    const syncFieldPreview = (field) => {
        if (!field) return;
        if (field.dataset.fieldType === "hardware_error") {
            syncHardwareErrorPreview(field);
        } else {
            syncRegularFieldPreview(field);
        }
    };
    const toggleValue = (input, option) => {
        const current = input.value.trim();
        const items = current ? current.split(",").map(s => s.trim()).filter(Boolean) : [];
        const idx = items.indexOf(option);
        if (idx === -1) {
            items.push(option);
        } else {
            items.splice(idx, 1);
        }
        input.value = items.join(", ");
    };
    const syncButtons = (input, buttons) => {
        const current = input.value.trim();
        const items = current ? current.split(",").map(s => s.trim()).filter(Boolean) : [];
        buttons.forEach(({ value, button }) => {
            button.classList.toggle("active", items.includes(value));
        });
    };
    const hasEmptyFields = () => {
        const fields = document.querySelectorAll(".fill-field");
        for (const field of fields) {
            const fieldType = field.dataset.fieldType;
            if (fieldType === "hardware_error") {
                const measurementInput = field.querySelector('[data-role="measurement_result"]');
                if (!measurementInput || measurementInput.value.trim() === "") {
                    return true;
                }
            } else {
                const control = field.querySelector('input:not(.hardware-error-input), select, textarea');
                if (control && control.value.trim() === "") {
                    return true;
                }
            }
        }
        return false;
    };
    document.querySelectorAll(".fill-field").forEach(syncFieldPreview);
    // Если есть пустые поля, сохранять при подтвеждении
    if (form) {
        form.addEventListener("submit", (event) => {
            if (allowSubmitWithEmptyFields) {
                allowSubmitWithEmptyFields = false;
                return;
            }
            if (hasEmptyFields()) {
                event.preventDefault();
                emptyFieldsModal?.show();
            }
        });
    }
    if (confirmSaveEmptyFieldsBtn && form) {
        confirmSaveEmptyFieldsBtn.addEventListener("click", () => {
            allowSubmitWithEmptyFields = true;
            emptyFieldsModal?.hide();
            form.requestSubmit();
        });
    }
    document.querySelectorAll(".fill-field").forEach(field => {
        field.addEventListener("input", () => syncFieldPreview(field));
        field.addEventListener("change", () => syncFieldPreview(field));
    });

    // Пользовательские типы для замен включают настраиваемые пользователем текстовые списки
    // Что позволяет значительно ускорить заполнение полей, требующих для ввода набора фиксированных значений
    fetch("/api/custom-types")
        .then(response => response.ok ? response.json() : [])
        .then(types => {
            const typeMap = new Map(
                (Array.isArray(types) ? types : []).map(item => [item.key, item])
            );
            document.querySelectorAll(".fill-field").forEach(field => {
                const typeKey = field.dataset.fieldType;
                const type = typeMap.get(typeKey);
                const input = field.querySelector('input[type="text"]:not(.hardware-error-input)');
                const box = field.querySelector(".custom-type-options");
                if (!type || !input || !box) return;
                box.classList.remove("d-none");
                const buttons = type.options.map(option => {
                    const button = document.createElement("button");
                    button.type = "button";
                    button.className = "btn btn-sm btn-outline-primary";
                    button.textContent = option;
                    button.addEventListener("click", () => {
                        toggleValue(input, option);
                        syncButtons(input, buttons);
                        syncFieldPreview(field);
                    });
                    box.appendChild(button);
                    return { value: option, button };
                });
                input.addEventListener("input", () => {
                    syncButtons(input, buttons);
                    syncFieldPreview(field);
                });
                syncButtons(input, buttons);
            });
        })
        .catch(console.error);
});
