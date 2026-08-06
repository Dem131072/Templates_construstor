/* Конструктор DOCX-шаблонов
 *
 * Этот файл отвечает только за браузерную часть: 
 * выбор фрагмента текста,
 * создание/редактирование плейсхолдера, 
 * управление пользовательскими типами,
 * рассчет позиций текста для замены
 * отправку итогового списка замен на сервер.
 */

(function () {
    'use strict';
    var BUILTIN_TYPES = [
        { value: 'text', label: 'Строка' },
        { value: 'date', label: 'Дата' },
        { value: 'time', label: 'Время' },
        { value: 'number', label: 'Число' },
        { value: 'select', label: 'Выбор' },
        { value: 'hardware_error', label: 'Аппаратная погрешность' }
    ];
    var TYPE_FORMATS = {
        date: [
            { value: '', label: 'Стандартный (dd.mm.yyyy)' },
            { value: 'full', label: '"01" ноября 2025 г.' }
        ],
        time: [
            { value: '', label: 'Стандартный (00:00)' },
            { value: 'hour_min', label: '00 час. 00 мин' }
        ],
    };

    function $(id) {
        return document.getElementById(id);
    }

    var els = {
        editor: $('editor'),
        form: $('save-form'),
        saveBtn: $('btn-save'),
        replacementsInput: $('replacements'),
        forceInput: $('force-input'),
        errorMessage: $('error-message'),
        errorList: $('error-list'),
        forceSaveBtn: $('force-save'),
        createBtn: $('create-field-btn'),
        editPlaceholderBtn: $('edit-placeholder-btn'),
        deletePlaceholderBtn: $('delete-placeholder-btn'),
        fieldName: $('field-name'),
        fieldType: $('field-type'),
        fieldFormat: $('field-format'),
        formatDiv: $('format-div'),
        saveFieldBtn: $('save-field'),
        customTypesBtn: $('custom-types-btn'),
        customTypesList: $('custom-types-list'),
        customTypesEmpty: $('custom-types-empty'),
        createCustomTypeBtn: $('create-custom-type-btn'),
        editCustomTypeBtn: $('edit-custom-type-btn'),
        deleteCustomTypeBtn: $('delete-custom-type-btn'),
        customTypeName: $('custom-type-name'),
        customTypeOptionInput: $('custom-type-option-input'),
        addCustomTypeOptionBtn: $('add-custom-type-option-btn'),
        customTypeOptionsList: $('custom-type-options-list'),
        removeCustomTypeOptionBtn: $('remove-custom-type-option-btn'),
        saveCustomTypeBtn: $('save-custom-type-btn'),
        folderSelect: $('folder-select'),
        targetFolderInput: $('target-folder-input'),
        confirmFolderBtn: $('confirm-folder-btn')
    };

    var state = {
        replacements: [],
        nextId: 0,
        currentRange: null,
        currentPlaceholder: null,
        isEditingPlaceholder: false,
        customTypes: [],
        selectedCustomTypeKey: '',
        editorTypeKey: '',
        editorOptions: [],
        selectedEditorOptionIndex: -1,
        returnToTypesModal: false
    };
    function getModal(id) {
        var el = $(id);
        if (!el || !window.bootstrap || !bootstrap.Modal) return null;
        return bootstrap.Modal.getOrCreateInstance(el);
    }
    function showModal(id) {
        var modal = getModal(id);
        if (modal) modal.show();
    }
    function hideModal(id) {
        var modal = getModal(id);
        if (modal) modal.hide();
    }
    function requestJson(url, options, fallbackMessage) {
        return fetch(url, options || {}).then(function (response) {
            if (!response.ok) {
                return response.text().then(function (text) {
                    throw new Error(text || fallbackMessage || 'Ошибка запроса');
                });
            }
            return response.json();
        });
    }
    function showRequestError(error, message) {
        console.error(error);
        alert(message);
    }
    function codePointLength(str) {
        return Array.from(str || '').length;
    }
    function normalizeSpaces(str) {
        return String(str || '')
            .replace(/\u00A0/g, ' ')
            .replace(/\u202F/g, ' ')
            .replace(/\u2007/g, ' ');
    }
    function normalizeTypeKey(str) {
        return String(str || '')
            .trim()
            .replace(/[\s:{};,"']/g, '_')
            .replace(/_+/g, '_')
            .replace(/^_+|_+$/g, '');
    }
    function parsePlaceholder(ph) {
        var parts = String(ph || '').split(':');
        return {
            name: parts[0] || '',
            fieldType: parts[1] || 'text',
            format: parts[2] || ''
        };
    }
    function buildPlaceholder(name, type, format) {
        return name + ':' + type + (format ? ':' + format : '');
    }
    function isBuiltInType(type) {
        for (var i = 0; i < BUILTIN_TYPES.length; i++) {
            if (BUILTIN_TYPES[i].value === type) return true;
        }
        return false;
    }
    function validatePlaceholder(type, format) {
        if (!type) return false;
        if (!isBuiltInType(type)) {
            return format === '';
        }
        if (type === 'text' || type === 'number' || type === 'select' || type === 'hardware_error') {
            return format === '';
        }
        if (type === 'date') {
            return format === '' || format === 'full';
        }
        if (type === 'time') {
            return format === '' || format === 'hour_min';
        }
        return false;
    }

    // Для отображения уже замененного пользователем текста
    function createPlaceholderSpan(data) {
        var span = document.createElement('span');
        span.className = 'placeholder';
        span.contentEditable = 'false';
        span.textContent = data.name;
        span.dataset.id = String(data.id);
        span.dataset.ph = data.ph;
        span.dataset.old = data.oldText;
        return span;
    }

    // Позиционирование и скрытие плавающих кнопок
    function hideActionButtons() {
        if (els.createBtn) els.createBtn.style.display = 'none';
        if (els.editPlaceholderBtn) els.editPlaceholderBtn.style.display = 'none';
        if (els.deletePlaceholderBtn) els.deletePlaceholderBtn.style.display = 'none';
    }
    function placeFloatingButton(button, x, y) {
        if (!button) return;
        button.style.left = x + 'px';
        button.style.top = Math.max(0, y - 40) + 'px';
        button.style.display = 'block';
    }
    function showCreateButton(x, y) {
        placeFloatingButton(els.createBtn, x, y);
        if (els.editPlaceholderBtn) els.editPlaceholderBtn.style.display = 'none';
        if (els.deletePlaceholderBtn) els.deletePlaceholderBtn.style.display = 'none';
    }
    function showPlaceholderButtons(x, y) {
        var gap = 12;
        placeFloatingButton(els.editPlaceholderBtn, x, y);
        var editWidth = els.editPlaceholderBtn ? els.editPlaceholderBtn.offsetWidth : 110;
        placeFloatingButton(els.deletePlaceholderBtn, x + editWidth + gap, y);
        if (els.createBtn) els.createBtn.style.display = 'none';
    }
    function getCaretRangeFromPointSafe(x, y) {
        if (document.caretPositionFromPoint) {
            var pos = document.caretPositionFromPoint(x, y);
            if (!pos) return null;
            var range = document.createRange();
            range.setStart(pos.offsetNode, pos.offset);
            range.collapse(true);
            return range;
        }
        if (document.caretRangeFromPoint) {
            return document.caretRangeFromPoint(x, y);
        }
        return null;
    }
    function getSelectedText() {
        if (!state.currentRange || !state.currentRange.cloneContents) return '';
        return normalizeSpaces(state.currentRange.cloneContents().textContent || '');
    }

    // Рассчет позиций для замены на сервере

    function getParagraphElement(index) {
        return document.querySelector('p[data-para-index="' + index + '"]');
    }
    function getParagraphIndex() {
        var container = null;
        if (state.currentRange && state.currentRange.commonAncestorContainer) {
            container = state.currentRange.commonAncestorContainer;
        } else if (state.currentPlaceholder) {
            container = state.currentPlaceholder;
        }
        if (!container) return null;
        var node = container.nodeType === Node.ELEMENT_NODE ? container : container.parentNode;
        if (!node || !node.closest) return null;
        var paragraph = node.closest('p[data-para-index]');
        if (!paragraph) return null;
        var index = parseInt(paragraph.dataset.paraIndex, 10);
        return isNaN(index) ? null : index;
    }
    function calculateOffset(range, paragraphIndex) {
        var paragraph = getParagraphElement(paragraphIndex);
        if (!paragraph || !range || !range.startContainer) return 0;
        var prefixRange = document.createRange();
        prefixRange.setStart(paragraph, 0);
        prefixRange.setEnd(range.startContainer, range.startOffset);
        return codePointLength(normalizeSpaces(prefixRange.cloneContents().textContent || ''));
    }
    function clearNode(node) {
        while (node.firstChild) {
            node.removeChild(node.firstChild);
        }
    }
    function fillSelect(select, items, selectedValue) {
        clearNode(select);
        for (var i = 0; i < items.length; i++) {
            var option = document.createElement('option');
            option.value = items[i].value;
            option.textContent = items[i].label;
            if (items[i].value === selectedValue) option.selected = true;
            select.appendChild(option);
        }
    }


    function renderFieldTypeOptions(selectedValue) {
        if (!els.fieldType) return;
        var current = selectedValue || els.fieldType.value || 'text';
        clearNode(els.fieldType);
        var builtinGroup = document.createElement('optgroup');
        builtinGroup.label = 'Стандартные типы';
        for (var i = 0; i < BUILTIN_TYPES.length; i++) {
            var opt = document.createElement('option');
            opt.value = BUILTIN_TYPES[i].value;
            opt.textContent = BUILTIN_TYPES[i].label;
            builtinGroup.appendChild(opt);
        }
        els.fieldType.appendChild(builtinGroup);
        if (state.customTypes.length) {
            var customGroup = document.createElement('optgroup');
            customGroup.label = 'Пользовательские типы';
            for (var j = 0; j < state.customTypes.length; j++) {
                var c = document.createElement('option');
                c.value = state.customTypes[j].key;
                c.textContent = state.customTypes[j].name;
                customGroup.appendChild(c);
            }
            els.fieldType.appendChild(customGroup);
        }
        if (!isBuiltInType(current)) {
            var exists = false;
            for (var k = 0; k < els.fieldType.options.length; k++) {
                if (els.fieldType.options[k].value === current) {
                    exists = true;
                    break;
                }
            }
            if (!exists) {
                var extra = document.createElement('option');
                extra.value = current;
                extra.textContent = current;
                els.fieldType.appendChild(extra);
            }
        }
        els.fieldType.value = current;
        updateFormatOptions();
    }
    function updateFormatOptions() {
        if (!els.fieldType || !els.fieldFormat || !els.formatDiv) return;
        var formats = TYPE_FORMATS[els.fieldType.value];
        if (!formats) {
            els.formatDiv.style.display = 'none';
            clearNode(els.fieldFormat);
            return;
        }
        els.formatDiv.style.display = 'block';
        fillSelect(els.fieldFormat, formats, els.fieldFormat.value || formats[0].value);
    }
    function openFieldModal(data) {
        data = data || {};
        if (!els.fieldName || !els.fieldType) return;
        els.fieldName.value = data.name || '';
        renderFieldTypeOptions(data.fieldType || 'text');
        if (TYPE_FORMATS[data.fieldType || 'text']) {
            els.fieldFormat.value = data.format || '';
        }
        showModal('fieldModal');
    }

    // Для загрузки в конструктор уже созданных шаблонов
    function initializeExistingPlaceholders() {
        if (!els.editor) return;
        state.replacements = [];
        state.nextId = 0;
        var spans = els.editor.querySelectorAll('.placeholder');
        for (var i = 0; i < spans.length; i++) {
            var span = spans[i];
            var id = state.nextId++;
            span.dataset.id = String(id);
            var paragraph = span.closest('p[data-para-index]');
            var paragraphIndex = 0;
            var offset = 0;
            if (paragraph) {
                paragraphIndex = parseInt(paragraph.dataset.paraIndex, 10) || 0;
                var range = document.createRange();
                range.setStart(paragraph, 0);
                range.setEndBefore(span);
                offset = codePointLength(range.toString());
            }
            var ph = span.dataset.ph || '';
            state.replacements.push({
                id: id,
                old: '{{' + ph + '}}',
                ph: ph,
                insert: '{{' + ph + '}}',
                paragraph_index: paragraphIndex,
                offset: offset
            });
        }
    }

    function savePlaceholder() {
        if (!els.fieldName || !els.fieldType) return;
        var name = (els.fieldName.value || '').trim();
        var type = els.fieldType.value;
        var format = TYPE_FORMATS[type] ? (els.fieldFormat.value || '') : '';
        if (!name) {
            alert('Название поля не может быть пустым');
            return;
        }
        if (!validatePlaceholder(type, format)) {
            alert('Выбран некорректный тип поля');
            return;
        }
        var ph = buildPlaceholder(name, type, format);
        if (state.isEditingPlaceholder && state.currentPlaceholder) {
            var editId = parseInt(state.currentPlaceholder.dataset.id, 10);
            for (var i = 0; i < state.replacements.length; i++) {
                if (state.replacements[i].id === editId) {
                    state.replacements[i].ph = ph;
                    state.replacements[i].insert = '{{' + ph + '}}';
                    break;
                }
            }
            state.currentPlaceholder.textContent = name;
            state.currentPlaceholder.dataset.ph = ph;
            state.currentPlaceholder = null;
            state.isEditingPlaceholder = false;
        } else {
            var oldText = getSelectedText();
            var paragraphIndex = getParagraphIndex();
            if (!oldText) {
                alert('Сначала выделите текст для замены');
                return;
            }
            // предполагается отсутствие необходимости в таких плейсхолдерах
            if (paragraphIndex === null) {
                alert('Текст для замены пересекает несколько параграфов.');
                return;
            }
            var id = state.nextId++;
            var offset = calculateOffset(state.currentRange, paragraphIndex);
            state.replacements.push({
                id: id,
                old: oldText,
                ph: ph,
                insert: '{{' + ph + '}}',
                paragraph_index: paragraphIndex,
                offset: offset
            });
            var span = createPlaceholderSpan({
                id: id,
                name: name,
                ph: ph,
                oldText: oldText
            });
            if (state.currentRange) {
                if (!state.currentRange.collapsed) {
                    state.currentRange.deleteContents();
                }
                state.currentRange.insertNode(span);
                state.currentRange.collapse(false);
            }
        }
        hideModal('fieldModal');
        els.fieldName.value = '';
    }
    function buildReplacementsPayload() {
        var byParagraph = new Map();
        var result = [];
        for (var i = 0; i < state.replacements.length; i++) {
            var rep = state.replacements[i];
            if (!byParagraph.has(rep.paragraph_index)) {
                byParagraph.set(rep.paragraph_index, []);
            }
            byParagraph.get(rep.paragraph_index).push(rep);
        }
        byParagraph.forEach(function (group) {
            group.sort(function (a, b) { return a.id - b.id; });
            var shifts = [];
            for (var i = 0; i < group.length; i++) {
                var rep = group[i];
                var shiftDelta = 0;
                for (var j = 0; j < shifts.length; j++) {
                    if (shifts[j].pos < rep.offset) {
                        shiftDelta += shifts[j].delta;
                    }
                }
                var originalOffset = rep.offset - shiftDelta;
                var nameLen = codePointLength((rep.ph.split(':')[0] || '').trim());
                var oldLen = codePointLength(rep.old);
                var delta = nameLen - oldLen;
                shifts.push({ pos: rep.offset, delta: delta });
                if (rep.old !== rep.insert) {
                    result.push({
                        old: rep.old,
                        insert: rep.insert,
                        paragraph_index: rep.paragraph_index,
                        offset: originalOffset
                    });
                }
            }
        });
        return result;
    }
    function showError(message) {
        if (!els.errorMessage) return;
        els.errorMessage.textContent = message;
        els.errorMessage.style.display = 'block';
    }
    function hideError() {
        if (!els.errorMessage) return;
        els.errorMessage.textContent = '';
        els.errorMessage.style.display = 'none';
    }

    // Сериализация replacements
    function handleSave(force) {
        if (!els.form || !els.replacementsInput || !els.forceInput) return;
        var invalid = [];
        var spans = els.editor ? els.editor.querySelectorAll('.placeholder') : [];
        for (var i = 0; i < spans.length; i++) {
            var info = parsePlaceholder(spans[i].dataset.ph || '');
            if (!validatePlaceholder(info.fieldType, info.format)) {
                invalid.push(spans[i].dataset.ph || '');
            }
        }
        if (!state.replacements.length) {
            alert('Нет замен для отправки.');
            return;
        }
        if (invalid.length) {
            showError('Некорректные поля: ' + invalid.join(', '));
            return;
        }
        var payload = buildReplacementsPayload();
        if (!payload.length && !state.replacements.length) {
            alert('Нет замен для отправки.');
            return;
        }
        hideError();
        els.replacementsInput.value = JSON.stringify(payload);
        els.forceInput.value = force ? 'true' : 'false';
        fetch('/constructor', {
            method: 'POST',
            body: new FormData(els.form)
        })
            .then(function (response) {
                if (response.redirected) {
                    window.location.href = response.url;
                    return null;
                }
                var contentType = response.headers.get('content-type') || '';
                if (contentType.indexOf('application/json') !== -1) {
                    return response.json().then(function (data) {
                        if (data && data.error && els.errorList) {
                            els.errorList.textContent = data.error;
                            showModal('errorModal');
                        }
                    });
                }
                if (contentType.indexOf('text/html') !== -1) {
                    return response.text().then(function (html) {
                        document.open();
                        document.write(html);
                        document.close();
                    });
                }
                alert('Неожиданный ответ сервера');
                return null;
            })
            .catch(function (error) {
                console.error(error);
                alert('Ошибка соединения с сервером');
            });
    }

    // Типы для замен, создаваемые пользователем
    function loadCustomTypes(showErrors) {
        return requestJson('/api/custom-types', {}, 'Не удалось загрузить пользовательские типы')
            .then(function (items) {
                state.customTypes = Array.isArray(items) ? items : [];
                renderFieldTypeOptions();
                renderCustomTypesList();
            })
            .catch(function (error) {
                console.error(error);
                if (showErrors) alert('Не удалось загрузить пользовательские типы');
            });
    }
    function renderCustomTypesList() {
        if (!els.customTypesList || !els.customTypesEmpty) return;
        clearNode(els.customTypesList);
        els.customTypesEmpty.style.display = state.customTypes.length ? 'none' : 'block';
        for (var i = 0; i < state.customTypes.length; i++) {
            var item = state.customTypes[i];
            var row = document.createElement('button');
            row.type = 'button';
            row.className = 'list-group-item list-group-item-action custom-type-row';
            if (state.selectedCustomTypeKey === item.key) {
                row.className += ' active';
            }
            row.dataset.key = item.key;
            var title = document.createElement('div');
            title.className = 'fw-semibold';
            title.textContent = item.name;
            var key = document.createElement('div');
            key.className = 'text-muted small';
            key.textContent = item.key;
            var options = document.createElement('div');
            options.className = 'small mt-1';
            options.textContent = Array.isArray(item.options) ? item.options.join(', ') : '';
            row.appendChild(title);
            row.appendChild(key);
            row.appendChild(options);
            els.customTypesList.appendChild(row);
        }
        var disabled = !state.selectedCustomTypeKey;
        if (els.editCustomTypeBtn) els.editCustomTypeBtn.disabled = disabled;
        if (els.deleteCustomTypeBtn) els.deleteCustomTypeBtn.disabled = disabled;
    }
    function resetCustomTypeEditor(type) {
        type = type || null;
        state.editorTypeKey = type ? type.key : '';
        state.editorOptions = type && Array.isArray(type.options) ? type.options.slice() : [];
        state.selectedEditorOptionIndex = -1;

        if (els.customTypeName) els.customTypeName.value = type ? (type.name || '') : '';
        if (els.customTypeOptionInput) els.customTypeOptionInput.value = '';

        renderCustomTypeEditorOptions();
    }
    function renderCustomTypeEditorOptions() {
        if (!els.customTypeOptionsList) return;
        clearNode(els.customTypeOptionsList);
        if (!state.editorOptions.length) {
            var empty = document.createElement('div');
            empty.className = 'list-group-item text-muted';
            empty.textContent = 'Варианты пока не добавлены';
            els.customTypeOptionsList.appendChild(empty);
        } else {
            for (var i = 0; i < state.editorOptions.length; i++) {
                var row = document.createElement('button');
                row.type = 'button';
                row.className = 'list-group-item list-group-item-action type-option-item';
                if (state.selectedEditorOptionIndex === i) {
                    row.className += ' active';
                }
                row.dataset.index = String(i);
                row.textContent = state.editorOptions[i];
                els.customTypeOptionsList.appendChild(row);
            }
        }
        if (els.removeCustomTypeOptionBtn) {
            els.removeCustomTypeOptionBtn.disabled = state.selectedEditorOptionIndex < 0;
        }
    }

    // Пользовательские типы в модальном окне
    function openCustomTypeEditor(type) {
        resetCustomTypeEditor(type);
        state.returnToTypesModal = true;
        hideModal('customTypesModal');
        setTimeout(function () {
            showModal('customTypeEditorModal');
        }, 120);
    }
    function addCustomTypeOption() {
        if (!els.customTypeOptionInput) return;
        var value = (els.customTypeOptionInput.value || '').trim();
        if (!value) return;
        if (state.editorOptions.indexOf(value) === -1) {
            state.editorOptions.push(value);
        }
        els.customTypeOptionInput.value = '';
        state.selectedEditorOptionIndex = -1;
        renderCustomTypeEditorOptions();
    }
    function saveCustomType() {
        var name = els.customTypeName ? (els.customTypeName.value || '').trim() : '';
        var key = normalizeTypeKey(name);
        var uniqueOptions = [];
        var i;
        if (!name) {
            alert('Название типа не может быть пустым');
            return;
        }
        for (i = 0; i < state.editorOptions.length; i++) {
            var item = (state.editorOptions[i] || '').trim();
            if (item && uniqueOptions.indexOf(item) === -1) {
                uniqueOptions.push(item);
            }
        }
        if (!uniqueOptions.length) {
            alert('Добавьте хотя бы один вариант текста');
            return;
        }
        var deletePromise = state.editorTypeKey && state.editorTypeKey !== key
            ? fetch('/api/custom-types/' + encodeURIComponent(state.editorTypeKey), { method: 'DELETE' })
            : Promise.resolve();
        deletePromise
            .then(function () {
                return requestJson('/api/custom-types', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ key: key, name: name, options: uniqueOptions })
                }, 'Не удалось сохранить тип');
            })
            .then(function () {
                state.selectedCustomTypeKey = key;
                state.returnToTypesModal = false;
                hideModal('customTypeEditorModal');
                return loadCustomTypes(false).then(function () {
                    setTimeout(function () { showModal('customTypesModal'); }, 120);
                });
            })
            .catch(function (error) {
                showRequestError(error, 'Не удалось сохранить пользовательский тип');
            });
    }
    function deleteSelectedCustomType() {
        if (!state.selectedCustomTypeKey) return;
        if (!window.confirm('Удалить выбранный пользовательский тип?')) return;
        fetch('/api/custom-types/' + encodeURIComponent(state.selectedCustomTypeKey), { method: 'DELETE' })
            .then(function (response) {
                if (!response.ok) throw new Error('Не удалось удалить тип');
                state.selectedCustomTypeKey = '';
                return loadCustomTypes(false);
            })
            .catch(function (error) {
                showRequestError(error, 'Не удалось удалить пользовательский тип');
            });
    }

    // Плавающие кнопки
    function handleEditorMouseUp(event) {
        var target = event.target;
        var targetElement = null;
        if (target) {
            targetElement = target.nodeType === Node.ELEMENT_NODE ? target : target.parentElement;
        }
        var placeholder = targetElement && targetElement.closest
            ? targetElement.closest('.placeholder')
            : null;
        if (placeholder && els.editor.contains(placeholder)) {
            state.currentPlaceholder = placeholder;
            state.currentRange = null;
            showPlaceholderButtons(event.pageX, event.pageY);
            return;
        }
        var selection = window.getSelection();
        if (selection && selection.rangeCount && selection.toString()) {
            state.currentPlaceholder = null;
            state.currentRange = selection.getRangeAt(0);
            var rect = state.currentRange.getBoundingClientRect();
            showCreateButton(rect.left + window.scrollX, rect.top + window.scrollY);
            return;
        }
        state.currentPlaceholder = null;
        hideActionButtons();
    }
    function handleEditorContextMenu(event) {
        event.preventDefault();
        var target = event.target;
        if (target && target.classList && target.classList.contains('placeholder')) {
            state.currentPlaceholder = target;
            showPlaceholderButtons(event.pageX, event.pageY);
            return;
        }
        var selection = window.getSelection();
        if (selection && selection.rangeCount) {
            state.currentRange = selection.getRangeAt(0);
        }
        if (!state.currentRange || state.currentRange.collapsed) {
            state.currentRange = getCaretRangeFromPointSafe(event.clientX, event.clientY);
        }
        if (state.currentRange && !state.currentRange.collapsed) {
            showCreateButton(event.pageX, event.pageY);
        } else {
            hideActionButtons();
        }
    }
    function isActionButtonTarget(target) {
        return (
            (els.createBtn && els.createBtn.contains(target)) ||
            (els.editPlaceholderBtn && els.editPlaceholderBtn.contains(target)) ||
            (els.deletePlaceholderBtn && els.deletePlaceholderBtn.contains(target))
        );
    }
    function handleDocumentMouseDown(event) {
        var target = event.target;
        if (els.editor && els.editor.contains(target)) return;
        if (isActionButtonTarget(target)) return;
        state.currentPlaceholder = null;
        hideActionButtons();
    }

    // Функции работы с плейсхолдерами
    function handleCreatePlaceholder() {
        var oldText = getSelectedText();
        if (!oldText) {
            hideActionButtons();
            return;
        }
        if (getParagraphIndex() === null) {
            alert('Текст для замены пересекает несколько параграфов.');
            hideActionButtons();
            return;
        }
        state.isEditingPlaceholder = false;
        openFieldModal({ name: normalizeSpaces(oldText), fieldType: 'text', format: '' });
        hideActionButtons();
    }
    function handleEditPlaceholder() {
        if (!state.currentPlaceholder) return;
        var info = parsePlaceholder(state.currentPlaceholder.dataset.ph || '');
        state.isEditingPlaceholder = true;
        openFieldModal({
            name: info.name,
            fieldType: info.fieldType,
            format: info.format
        });
        hideActionButtons();
    }
    function handleDeletePlaceholder() {
        if (!state.currentPlaceholder) return;
        var id = parseInt(state.currentPlaceholder.dataset.id, 10);
        var oldText = state.currentPlaceholder.dataset.old || state.currentPlaceholder.textContent;
        var i;
        for (i = 0; i < state.replacements.length; i++) {
            if (state.replacements[i].id === id) {
                if ((state.replacements[i].old || '').indexOf('{{') === 0) {
                    state.replacements[i].insert = oldText;
                } else {
                    state.replacements.splice(i, 1);
                }
                break;
            }
        }
        state.currentPlaceholder.replaceWith(document.createTextNode(oldText));
        state.currentPlaceholder = null;
        hideActionButtons();
    }

    // Все DOM-обработчики
    function bindEvents() {
        if (els.editor) {
            els.editor.addEventListener('mouseup', handleEditorMouseUp);
            els.editor.addEventListener('mousedown', handleDocumentMouseDown);
            els.editor.addEventListener('contextmenu', handleEditorContextMenu);
        }
        if (els.createBtn) els.createBtn.addEventListener('click', handleCreatePlaceholder);
        if (els.editPlaceholderBtn) els.editPlaceholderBtn.addEventListener('click', handleEditPlaceholder);
        if (els.deletePlaceholderBtn) els.deletePlaceholderBtn.addEventListener('click', handleDeletePlaceholder);
        if (els.fieldType) els.fieldType.addEventListener('change', updateFormatOptions);
        if (els.saveFieldBtn) els.saveFieldBtn.addEventListener('click', savePlaceholder);
        if (els.saveBtn) {
            els.saveBtn.addEventListener('click', function (e) {
                e.preventDefault();
                requestJson('/api/template-folders', {}, 'Не удалось загрузить список папок')
                    .then(function (folders) {
                        clearNode(els.folderSelect);
                        folders.forEach(function (folder) {
                            var opt = document.createElement('option');
                            opt.value = folder;
                            opt.textContent = folder === '' ? 'doc_templates (корень)' : folder;
                            els.folderSelect.appendChild(opt);
                        });
                        showModal('folderSelectModal');
                    })
                    .catch(function (error) {
                        showRequestError(error, 'Не удалось загрузить список папок');
                    });
            });
        }
        if (els.confirmFolderBtn) {
            els.confirmFolderBtn.addEventListener('click', function () {
                els.targetFolderInput.value = els.folderSelect.value;
                hideModal('folderSelectModal');
                handleSave(false);
            });
        }
        if (els.form) {
            els.form.addEventListener('submit', function (e) {
                e.preventDefault();
            });
        }
        if (els.forceSaveBtn) {
            els.forceSaveBtn.addEventListener('click', function (event) {
                event.preventDefault();
                handleSave(true);
            });
        }
        if (els.customTypesBtn) {
            els.customTypesBtn.addEventListener('click', function () {
                loadCustomTypes(true).then(function () {
                    showModal('customTypesModal');
                });
            });
        }
        if (els.createCustomTypeBtn) {
            els.createCustomTypeBtn.addEventListener('click', function () {
                openCustomTypeEditor(null);
            });
        }
        if (els.editCustomTypeBtn) {
            els.editCustomTypeBtn.addEventListener('click', function () {
                var type = null;
                for (var i = 0; i < state.customTypes.length; i++) {
                    if (state.customTypes[i].key === state.selectedCustomTypeKey) {
                        type = state.customTypes[i];
                        break;
                    }
                }
                if (type) openCustomTypeEditor(type);
            });
        }
        if (els.deleteCustomTypeBtn) {
            els.deleteCustomTypeBtn.addEventListener('click', deleteSelectedCustomType);
        }
        if (els.addCustomTypeOptionBtn) {
            els.addCustomTypeOptionBtn.addEventListener('click', addCustomTypeOption);
        }
        if (els.removeCustomTypeOptionBtn) {
            els.removeCustomTypeOptionBtn.addEventListener('click', function () {
                if (state.selectedEditorOptionIndex < 0) return;
                state.editorOptions.splice(state.selectedEditorOptionIndex, 1);
                state.selectedEditorOptionIndex = -1;
                renderCustomTypeEditorOptions();
            });
        }
        if (els.saveCustomTypeBtn) {
            els.saveCustomTypeBtn.addEventListener('click', saveCustomType);
        }
        if (els.customTypesList) {
            els.customTypesList.addEventListener('click', function (event) {
                var row = event.target.closest('.custom-type-row');
                if (!row) return;
                state.selectedCustomTypeKey = row.dataset.key || '';
                renderCustomTypesList();
            });
        }
        if (els.customTypeOptionsList) {
            els.customTypeOptionsList.addEventListener('click', function (event) {
                var row = event.target.closest('.type-option-item');
                if (!row) return;
                state.selectedEditorOptionIndex = parseInt(row.dataset.index, 10);
                renderCustomTypeEditorOptions();
            });
        }
        if (els.customTypeOptionInput) {
            els.customTypeOptionInput.addEventListener('keydown', function (event) {
                if (event.key === 'Enter') {
                    event.preventDefault();
                    addCustomTypeOption();
                }
            });
        }
        document.addEventListener('keydown', function (event) {
            var active = document.activeElement;
            var fieldModalEl = $('fieldModal');
            if (
                event.key === 'Enter' &&
                fieldModalEl &&
                fieldModalEl.classList.contains('show') &&
                active &&
                (active.tagName === 'INPUT' || active.tagName === 'SELECT')
            ) {
                event.preventDefault();
                savePlaceholder();
            }
        });
        var editorModalEl = $('customTypeEditorModal');
        if (editorModalEl) {
            editorModalEl.addEventListener('hidden.bs.modal', function () {
                if (!state.returnToTypesModal) return;
                state.returnToTypesModal = false;
                setTimeout(function () {
                    showModal('customTypesModal');
                }, 120);
            });
        }
        window.addEventListener('load', initializeExistingPlaceholders);
    }
    bindEvents();
    loadCustomTypes(false);
})();
