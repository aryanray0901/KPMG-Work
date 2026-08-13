(() => {
  const app = document.getElementById('editor-app');
  if (!app) return;

  const sid = app.dataset.sid;
  let state = { ...window.DECK_EDITOR_STATE };
  let selectedSlide = state.selected_slide || 1;
  let busy = false;
  let previewRecoveryTimer = null;
  let previewRecoveryAttempts = 0;

  const thumbnailList = document.getElementById('thumbnail-list');
  const slideImage = document.getElementById('editor-slide-image');
  const loading = document.getElementById('editor-loading');
  const previewUnavailable = document.getElementById('preview-unavailable');
  const previewHelp = document.getElementById('preview-help');
  const rendererLabel = document.getElementById('renderer-label');
  const chatMessages = document.getElementById('chat-messages');
  const chatInput = document.getElementById('chat-input');
  const chatForm = document.getElementById('chat-form');
  const sendButton = document.getElementById('send-chat');
  const zoomSlider = document.getElementById('zoom-slider');
  const zoomLabel = document.getElementById('zoom-label');
  const attachmentInput = document.getElementById('chat-attachments');
  const attachmentLabel = document.getElementById('attachment-label');
  const themePreset = document.getElementById('theme-preset');
  const applyThemeSlideButton = document.getElementById('apply-theme-slide');
  const applyThemeDeckButton = document.getElementById('apply-theme-deck');
  const customThemeColors = document.getElementById('custom-theme-colors');
  const themePrimary = document.getElementById('theme-primary');
  const themeAccent = document.getElementById('theme-accent');
  const themeBackground = document.getElementById('theme-background');
  const builderModal = document.getElementById('builder-modal');
  const builderForm = document.getElementById('builder-form');
  const builderSubmit = builderForm.querySelector('button[type="submit"]');
  const builderLayout = document.getElementById('builder-layout');
  const builderChartType = document.getElementById('builder-chart-type');
  const chartDataFile = document.getElementById('chart-data-file');
  const chartDataControl = document.getElementById('chart-data-control');
  const chartDataStatus = document.getElementById('chart-data-status');
  const previousButton = document.getElementById('previous-slide');
  const nextButton = document.getElementById('next-slide');
  const guidedModal = document.getElementById('guided-modal');
  const guidedForm = document.getElementById('guided-form');
  const guidedFields = document.getElementById('guided-fields');
  const guidedTitle = document.getElementById('guided-title');
  const guidedDescription = document.getElementById('guided-description');
  const guidedCommandPreview = document.getElementById('guided-command-preview');
  const rightPanelTabs = Array.from(document.querySelectorAll('[data-right-tab]'));
  const rightPanelPanes = Array.from(document.querySelectorAll('[data-right-pane]'));
  let imageReplacePending = false;
  let activeGuidedAction = null;

  const guidedActions = {
    sort_table: {
      title: 'Sort table',
      description: 'Enter a column number or header name.',
      fields: [
        { name: 'column', label: 'Column number or name', value: '1', placeholder: '3 or Impact' },
        { name: 'direction', label: 'Direction', type: 'select', options: [['descending', 'Descending'], ['ascending', 'Ascending']] },
      ],
      command: (values) => `Sort the table on slide ${selectedSlide} by column "${cleanValue(values.column)}" in ${values.direction} order.`,
    },
    add_table_row: {
      title: 'Add table row',
      description: 'Separate each cell with a vertical bar.',
      fields: [{ name: 'values', label: 'Row values', value: 'New risk | Owner | Medium', placeholder: 'Risk | Owner | Impact' }],
      command: (values) => `Add a table row with values "${cleanValue(values.values)}" on slide ${selectedSlide}.`,
    },
    merge_table_cells: {
      title: 'Merge table cells',
      description: 'Enter the row and the two columns to merge.',
      fields: [
        { name: 'row', label: 'Row', type: 'number', value: '1', min: '1' },
        { name: 'first_column', label: 'First column', type: 'number', value: '1', min: '1' },
        { name: 'second_column', label: 'Second column', type: 'number', value: '2', min: '1' },
      ],
      command: (values) => `Merge columns ${values.first_column} and ${values.second_column} in row ${values.row} of the table on slide ${selectedSlide}.`,
    },
    split_table_cell: {
      title: 'Split table cell',
      description: 'Enter the row and column of the merged cell.',
      fields: [
        { name: 'row', label: 'Row', type: 'number', value: '1', min: '1' },
        { name: 'column', label: 'Column', type: 'number', value: '1', min: '1' },
      ],
      command: (values) => `Split the table cell in row ${values.row} and column ${values.column} on slide ${selectedSlide}.`,
    },
    set_table_cell: {
      title: 'Edit table cell',
      description: 'Choose the cell and enter its new text.',
      fields: [
        { name: 'row', label: 'Row', type: 'number', value: '2', min: '1' },
        { name: 'column', label: 'Column', type: 'number', value: '1', min: '1' },
        { name: 'text', label: 'New cell text', placeholder: 'High risk' },
      ],
      command: (values) => `Set the table cell in row ${values.row} and column ${values.column} to "${cleanValue(values.text)}" on slide ${selectedSlide}.`,
    },
    delete_table_row: {
      title: 'Delete table row',
      description: 'Enter the row number to remove.',
      fields: [{ name: 'row', label: 'Row', type: 'number', value: '2', min: '1' }],
      command: (values) => `Delete row ${values.row} from the table on slide ${selectedSlide}.`,
    },
    add_table_column: {
      title: 'Add table column',
      description: 'Start with the header, then separate each row value with a vertical bar.',
      fields: [{ name: 'values', label: 'Column values', value: 'Status | Open | Closed', placeholder: 'Header | Row 2 | Row 3' }],
      command: (values) => `Add a table column with values "${cleanValue(values.values)}" on slide ${selectedSlide}.`,
    },
    delete_table_column: {
      title: 'Delete table column',
      description: 'Enter the column number to remove.',
      fields: [{ name: 'column', label: 'Column', type: 'number', value: '2', min: '1' }],
      command: (values) => `Delete column ${values.column} from the table on slide ${selectedSlide}.`,
    },
    replace_text: {
      title: 'Replace text',
      description: 'Enter the exact text to find and its replacement.',
      fields: [
        { name: 'old_text', label: 'Find', placeholder: 'Q3' },
        { name: 'new_text', label: 'Replace with', placeholder: 'Q4' },
        { name: 'scope', label: 'Scope', type: 'select', options: [['slide', 'Current slide'], ['deck', 'Whole deck']] },
      ],
      command: (values) => `Replace "${cleanValue(values.old_text)}" with "${cleanValue(values.new_text)}" ${values.scope === 'deck' ? 'across the entire deck' : `on slide ${selectedSlide}`}.`,
    },
    merge_slides: {
      title: 'Merge slides',
      description: 'Enter the two slide numbers to combine.',
      fields: [
        { name: 'first_slide', label: 'First slide', type: 'number', value: () => String(selectedSlide), min: '1' },
        { name: 'second_slide', label: 'Second slide', type: 'number', value: () => String(Math.min(selectedSlide + 1, state.slide_count)), min: '1' },
      ],
      command: (values) => `Merge slides ${values.first_slide} and ${values.second_slide}.`,
    },
    move_slide: {
      title: 'Move slide',
      description: 'Enter the current and destination slide numbers.',
      fields: [
        { name: 'from_slide', label: 'Move slide', type: 'number', value: () => String(selectedSlide), min: '1' },
        { name: 'to_slide', label: 'To position', type: 'number', value: () => String(selectedSlide), min: '1' },
      ],
      command: (values) => `Move slide ${values.from_slide} to position ${values.to_slide}.`,
    },
  };

  function cleanValue(value) {
    return String(value || '').replace(/["“”]/g, "'").replace(/\s+/g, ' ').trim();
  }

  function activateRightTab(name, focusTab = false) {
    rightPanelTabs.forEach((tab) => {
      const active = tab.dataset.rightTab === name;
      tab.classList.toggle('active', active);
      tab.setAttribute('aria-selected', active ? 'true' : 'false');
      tab.tabIndex = active ? 0 : -1;
      if (active && focusTab) tab.focus();
    });
    rightPanelPanes.forEach((pane) => {
      pane.hidden = pane.dataset.rightPane !== name;
    });
  }

  rightPanelTabs.forEach((tab, index) => {
    tab.addEventListener('click', () => activateRightTab(tab.dataset.rightTab));
    tab.addEventListener('keydown', (event) => {
      let destination = null;
      if (event.key === 'ArrowRight') destination = (index + 1) % rightPanelTabs.length;
      if (event.key === 'ArrowLeft') destination = (index - 1 + rightPanelTabs.length) % rightPanelTabs.length;
      if (event.key === 'Home') destination = 0;
      if (event.key === 'End') destination = rightPanelTabs.length - 1;
      if (destination === null) return;
      event.preventDefault();
      activateRightTab(rightPanelTabs[destination].dataset.rightTab, true);
    });
  });

  function imageUrl(slide) {
    const revision = state.preview_revision || state.version || Date.now();
    return `/editor_slide_image/${sid}/${state.version}/${slide}?v=${revision}-${Date.now()}`;
  }

  function clearPreviewRecovery() {
    if (previewRecoveryTimer) window.clearTimeout(previewRecoveryTimer);
    previewRecoveryTimer = null;
    previewRecoveryAttempts = 0;
  }

  function schedulePreviewRecovery() {
    if (state.rendering_ok || previewRecoveryTimer || previewRecoveryAttempts >= 3) return;
    const delays = [900, 2200, 4500];
    const delay = delays[previewRecoveryAttempts] || 4500;
    previewRecoveryAttempts += 1;
    previewRecoveryTimer = window.setTimeout(async () => {
      previewRecoveryTimer = null;
      try {
        const data = await postJson(`/editor/action/${sid}`, {
          action: 'retry_preview',
          selected_slide: selectedSlide,
        });
        state = data.state;
        selectedSlide = state.selected_slide;
        if (state.rendering_ok) clearPreviewRecovery();
        renderAll();
      } catch (_) {
        schedulePreviewRecovery();
      }
    }, delay);
  }

  function clampSelected() {
    selectedSlide = Math.max(1, Math.min(selectedSlide, Math.max(1, state.slide_count)));
  }

  function renderThumbnails() {
    thumbnailList.innerHTML = '';
    for (let slide = 1; slide <= state.slide_count; slide += 1) {
      const button = document.createElement('button');
      button.type = 'button';
      button.className = `editor-thumb${slide === selectedSlide ? ' active' : ''}`;
      button.dataset.slide = String(slide);
      button.innerHTML = `<span>${slide}</span>${state.rendering_ok ? `<img src="${imageUrl(slide)}" alt="Slide ${slide}">` : '<div class="thumb-placeholder">Preview</div>'}`;
      button.addEventListener('click', () => selectSlide(slide));
      thumbnailList.appendChild(button);
    }
  }

  function updateButtons() {
    document.getElementById('undo-btn').disabled = busy || !state.can_undo;
    document.getElementById('redo-btn').disabled = busy || !state.can_redo;
    document.querySelectorAll('[data-action]').forEach((button) => {
      if (['undo', 'redo'].includes(button.dataset.action)) return;
      const atFirst = button.dataset.action === 'move_left' && selectedSlide <= 1;
      const atLast = button.dataset.action === 'move_right' && selectedSlide >= state.slide_count;
      button.disabled = busy || atFirst || atLast;
    });
    document.querySelectorAll('[data-prompt], [data-editor-form], [data-chart-type]').forEach((control) => {
      if ('disabled' in control) control.disabled = busy;
      control.classList.toggle('disabled', busy);
    });
    previousButton.disabled = busy || selectedSlide <= 1;
    nextButton.disabled = busy || selectedSlide >= state.slide_count;
    sendButton.disabled = busy || !chatInput.value.trim();
    if (applyThemeSlideButton) applyThemeSlideButton.disabled = busy;
    if (applyThemeDeckButton) applyThemeDeckButton.disabled = busy;
    updateBuilderControls();
  }

  function renderCurrentSlide() {
    clampSelected();
    document.getElementById('slide-label').textContent = `Slide ${selectedSlide} of ${state.slide_count}`;
    document.getElementById('slide-total').textContent = state.slide_count;
    document.getElementById('selected-slide-note').textContent = `Editing slide ${selectedSlide}`;
    document.getElementById('save-status').textContent = `Version ${state.version + 1}`;
    if (rendererLabel) rendererLabel.textContent = state.render_engine || 'Preview renderer unavailable';
    if (previewHelp) previewHelp.textContent = state.render_help || 'Retry the preview after installing a supported renderer.';
    if (previewUnavailable) previewUnavailable.hidden = Boolean(state.rendering_ok);
    if (slideImage) {
      slideImage.hidden = !state.rendering_ok;
      if (state.rendering_ok) {
        slideImage.classList.add('is-loading');
        slideImage.onload = () => slideImage.classList.remove('is-loading');
        slideImage.onerror = () => {
          slideImage.classList.remove('is-loading');
          slideImage.hidden = true;
          state.rendering_ok = false;
          if (previewUnavailable) previewUnavailable.hidden = false;
          schedulePreviewRecovery();
        };
        slideImage.src = imageUrl(selectedSlide);
      }
    }
    document.querySelectorAll('.editor-thumb').forEach((button) => {
      button.classList.toggle('active', Number(button.dataset.slide) === selectedSlide);
    });
  }

  function renderAll() {
    clampSelected();
    renderThumbnails();
    renderCurrentSlide();
    updateButtons();
    if (state.rendering_ok) clearPreviewRecovery();
    else schedulePreviewRecovery();
  }

  function selectSlide(slide) {
    selectedSlide = slide;
    renderCurrentSlide();
    updateButtons();
  }

  function addMessage(role, text, temporary = false) {
    const bubble = document.createElement('div');
    bubble.className = `chat-bubble ${role}${temporary ? ' temporary' : ''}`;
    bubble.textContent = text;
    chatMessages.appendChild(bubble);
    chatMessages.scrollTop = chatMessages.scrollHeight;
    return bubble;
  }

  function setBusy(value) {
    busy = value;
    if (loading) loading.hidden = !value;
    updateButtons();
  }

  async function postJson(url, payload) {
    const response = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    });
    let data;
    try {
      data = await response.json();
    } catch (_) {
      throw new Error('The server returned an invalid response.');
    }
    if (!response.ok || !data.ok) throw new Error(data.error || 'The edit failed.');
    return data;
  }


  async function postForm(url, formData) {
    const response = await fetch(url, { method: 'POST', body: formData });
    let data;
    try {
      data = await response.json();
    } catch (_) {
      throw new Error('The server returned an invalid response.');
    }
    if (!response.ok || !data.ok) throw new Error(data.error || 'The edit could not be completed.');
    return data;
  }

  async function runAction(action) {
    if (busy) return;
    setBusy(true);
    try {
      const data = await postJson(`/editor/action/${sid}`, { action, selected_slide: selectedSlide });
      state = data.state;
      selectedSlide = state.selected_slide;
      previewRecoveryAttempts = 0;
      addMessage('assistant', data.message);
      renderAll();
    } catch (error) {
      addMessage('assistant error', error.message);
    } finally {
      setBusy(false);
    }
  }

  document.querySelectorAll('[data-action]').forEach((button) => {
    button.addEventListener('click', () => runAction(button.dataset.action));
  });

  const chartLayoutByType = {
    column: 'bar_chart', bar: 'bar_chart', line: 'line_chart', pie: 'pie_chart',
    area: 'area_chart', waterfall: 'waterfall_chart', scatter: 'scatter_plot',
  };

  function selectedDataMode() {
    return builderForm.querySelector('input[name="data_mode"]:checked')?.value || 'blank';
  }

  function selectedPlacement() {
    return builderForm.querySelector('input[name="placement"]:checked')?.value || 'new';
  }

  function updateBuilderControls() {
    const uploadMode = selectedDataMode() === 'upload';
    const file = chartDataFile.files[0];
    chartDataControl.hidden = !uploadMode;
    chartDataStatus.textContent = file ? file.name : 'Choose one worksheet file.';
    builderSubmit.disabled = busy || (uploadMode && !file);
    builderSubmit.textContent = selectedPlacement() === 'current' ? 'Add to current slide' : 'Create new slide';
  }

  function openChartBuilder(chartType = 'column', dataMode = 'blank') {
    const chosenType = chartLayoutByType[chartType] ? chartType : 'column';
    builderChartType.value = chosenType;
    builderLayout.value = chartLayoutByType[chosenType];
    const mode = builderForm.querySelector(`input[name="data_mode"][value="${dataMode === 'upload' ? 'upload' : 'blank'}"]`);
    if (mode) mode.checked = true;
    const placement = builderForm.querySelector('input[name="placement"][value="new"]');
    if (placement) placement.checked = true;
    builderModal.hidden = false;
    updateBuilderControls();
    if (dataMode === 'upload') window.setTimeout(() => chartDataFile.focus(), 0);
  }

  function closeBuilder() { builderModal.hidden = true; }
  document.getElementById('new-slide-button').addEventListener('click', () => openChartBuilder('column', 'blank'));
  document.querySelectorAll('[data-chart-type]').forEach((button) => {
    button.addEventListener('click', () => openChartBuilder(button.dataset.chartType, 'blank'));
  });
  document.getElementById('builder-close').addEventListener('click', closeBuilder);
  document.getElementById('builder-cancel').addEventListener('click', closeBuilder);
  builderModal.addEventListener('click', (event) => { if (event.target === builderModal) closeBuilder(); });
  builderChartType.addEventListener('change', () => {
    builderLayout.value = chartLayoutByType[builderChartType.value] || 'bar_chart';
  });
  builderForm.querySelectorAll('input[name="data_mode"]').forEach((control) => {
    control.addEventListener('change', updateBuilderControls);
  });
  builderForm.querySelectorAll('input[name="placement"]').forEach((control) => {
    control.addEventListener('change', updateBuilderControls);
  });
  chartDataFile.addEventListener('change', () => {
    updateBuilderControls();
  });

  builderForm.addEventListener('submit', async (event) => {
    event.preventDefault();
    if (busy) return;
    const chartType = builderChartType.value;
    const uploadMode = selectedDataMode() === 'upload';
    const placement = selectedPlacement();
    const file = chartDataFile.files[0];
    if (uploadMode && !file) {
      updateBuilderControls();
      return;
    }
    const formData = new FormData(builderForm);
    formData.set('layout', chartLayoutByType[chartType] || 'bar_chart');
    formData.set('chart_type', chartType);
    formData.set('selected_slide', String(selectedSlide));
    formData.set('placement', placement);
    formData.set('smart', 'false');
    if (uploadMode && file) formData.set('data_file', file);
    closeBuilder();
    const destination = placement === 'current' ? `on slide ${selectedSlide}` : 'on a new slide';
    const thinking = addMessage('assistant', `Adding a native editable ${chartType} chart ${destination}…`, true);
    setBusy(true);
    try {
      const data = await postForm(`/editor/build/${sid}`, formData);
      thinking.remove();
      addMessage('assistant', data.message);
      state = data.state;
      selectedSlide = state.selected_slide;
      previewRecoveryAttempts = 0;
      renderAll();
      builderForm.reset();
      builderChartType.value = 'column';
      builderLayout.value = 'bar_chart';
      chartDataFile.value = '';
      updateBuilderControls();
    } catch (error) {
      thinking.remove();
      addMessage('assistant error', error.message);
    } finally {
      setBusy(false);
    }
  });

  async function runQuickCommand(command) {
      if (busy || !command) return;
      addMessage('user', command);
      const thinking = addMessage('assistant', 'Applying the native PowerPoint edit…', true);
      setBusy(true);
      try {
        const data = await postJson(`/editor/quick/${sid}`, { command, selected_slide: selectedSlide });
        thinking.remove();
        addMessage('assistant', data.message);
        state = data.state;
        selectedSlide = state.selected_slide;
        previewRecoveryAttempts = 0;
        renderAll();
      } catch (error) {
        thinking.remove();
        addMessage('assistant error', error.message);
      } finally {
        setBusy(false);
      }
  }

  function guidedValues() {
    return Object.fromEntries(Array.from(guidedForm.elements)
      .filter((element) => element.name)
      .map((element) => [element.name, element.value]));
  }

  function updateGuidedPreview() {
    if (!activeGuidedAction) return;
    guidedCommandPreview.textContent = activeGuidedAction.command(guidedValues());
  }

  function openGuided(actionName, defaults = {}) {
    const action = guidedActions[actionName];
    if (!action || busy) return;
    activeGuidedAction = action;
    guidedTitle.textContent = action.title;
    guidedDescription.textContent = action.description;
    guidedFields.innerHTML = '';
    action.fields.forEach((field) => {
      const label = document.createElement('label');
      label.textContent = field.label;
      let control;
      if (field.type === 'select') {
        control = document.createElement('select');
        field.options.forEach(([value, text]) => {
          const option = document.createElement('option');
          option.value = value;
          option.textContent = text;
          control.appendChild(option);
        });
      } else {
        control = document.createElement('input');
        control.type = field.type || 'text';
        if (field.placeholder) control.placeholder = field.placeholder;
        if (field.min) control.min = field.min;
        if (field.max) control.max = field.max;
      }
      control.name = field.name;
      control.required = field.required !== false;
      const baseValue = typeof field.value === 'function' ? field.value() : field.value;
      control.value = defaults[field.name] || baseValue || '';
      label.appendChild(control);
      guidedFields.appendChild(label);
    });
    guidedModal.hidden = false;
    updateGuidedPreview();
    const first = guidedFields.querySelector('input, select');
    if (first) first.focus();
  }

  function closeGuided() {
    guidedModal.hidden = true;
    activeGuidedAction = null;
    guidedFields.innerHTML = '';
  }

  document.querySelectorAll('[data-editor-form]').forEach((button) => {
    button.addEventListener('click', () => openGuided(button.dataset.editorForm));
  });
  guidedForm.addEventListener('input', updateGuidedPreview);
  guidedForm.addEventListener('submit', (event) => {
    event.preventDefault();
    if (!activeGuidedAction || !guidedForm.reportValidity()) return;
    const command = activeGuidedAction.command(guidedValues());
    closeGuided();
    runQuickCommand(command);
  });
  document.getElementById('guided-close').addEventListener('click', closeGuided);
  document.getElementById('guided-cancel').addEventListener('click', closeGuided);
  guidedModal.addEventListener('click', (event) => { if (event.target === guidedModal) closeGuided(); });

  document.querySelectorAll('[data-prompt]').forEach((button) => {
    button.addEventListener('click', () => {
      if (button.hasAttribute('data-image-replace')) return;
      runQuickCommand(button.dataset.prompt);
    });
  });

  document.querySelectorAll('[data-image-replace]').forEach((button) => {
    button.addEventListener('click', () => {
      if (busy || !attachmentInput) return;
      imageReplacePending = true;
      attachmentInput.value = '';
      attachmentInput.click();
    });
  });

  const inspectorToggle = document.getElementById('inspector-toggle');
  if (inspectorToggle) {
    inspectorToggle.addEventListener('click', () => {
      const body = document.getElementById('inspector-body');
      body.hidden = !body.hidden;
      inspectorToggle.textContent = body.hidden ? 'Show' : 'Hide';
    });
  }

  previousButton.addEventListener('click', () => selectSlide(selectedSlide - 1));
  nextButton.addEventListener('click', () => selectSlide(selectedSlide + 1));

  chatForm.addEventListener('submit', async (event) => {
    event.preventDefault();
    const message = chatInput.value.trim();
    if (!message || busy) return;
    addMessage('user', message);
    chatInput.value = '';
    const thinking = addMessage('assistant', 'Analyzing every slide, choosing the best edits, and updating the PowerPoint…', true);
    setBusy(true);
    try {
      const formData = new FormData();
      formData.append('message', message);
      formData.append('selected_slide', String(selectedSlide));
      if (attachmentInput) {
        Array.from(attachmentInput.files || []).slice(0, 10).forEach((file) => formData.append('attachments', file));
      }
      const data = await postForm(`/editor/chat/${sid}`, formData);
      thinking.remove();
      addMessage('assistant', data.message);
      if (attachmentInput) attachmentInput.value = '';
      if (attachmentLabel) attachmentLabel.textContent = 'No attachments';
      state = data.state;
      selectedSlide = state.selected_slide;
      previewRecoveryAttempts = 0;
      renderAll();
    } catch (error) {
      thinking.remove();
      addMessage('assistant error', error.message);
    } finally {
      setBusy(false);
      chatInput.focus();
    }
  });

  chatInput.addEventListener('keydown', (event) => {
    if (event.key === 'Enter' && (event.metaKey || event.ctrlKey)) chatForm.requestSubmit();
  });
  chatInput.addEventListener('input', updateButtons);


  if (attachmentInput) {
    attachmentInput.addEventListener('change', () => {
      const files = Array.from(attachmentInput.files || []);
      attachmentLabel.textContent = files.length ? files.map((file) => file.name).join(', ') : 'No attachments';
      if (imageReplacePending) {
        imageReplacePending = false;
        if (files.length) {
          chatInput.value = 'Replace the main image on this slide using my attached image.';
          chatForm.requestSubmit();
        }
      }
    });
  }


  function applySelectedTheme(scope) {
    if (busy) return;
    const scopeText = scope === 'slide' ? `slide ${selectedSlide}` : 'the entire deck';
    if (themePreset && themePreset.value === 'Custom') {
      const primary = themePrimary ? themePrimary.value : '#00338D';
      const accent = themeAccent ? themeAccent.value : '#00A651';
      const background = themeBackground ? themeBackground.value : '#F7F9FC';
      runQuickCommand(`Apply a custom theme to ${scopeText} with primary ${primary}, accent ${accent}, and background ${background}. Preserve logos and status colors.`);
      return;
    }
    const preset = themePreset ? themePreset.value : 'Deck Refresh Blue';
    runQuickCommand(`Apply the ${preset} theme to ${scopeText}. Preserve logos and status colors.`);
  }

  if (themePreset) {
    themePreset.addEventListener('change', () => {
      customThemeColors.hidden = themePreset.value !== 'Custom';
    });
  }
  if (applyThemeSlideButton) applyThemeSlideButton.addEventListener('click', () => applySelectedTheme('slide'));
  if (applyThemeDeckButton) applyThemeDeckButton.addEventListener('click', () => applySelectedTheme('deck'));

  zoomSlider.addEventListener('input', () => {
    const zoom = Number(zoomSlider.value);
    zoomLabel.textContent = `${zoom}%`;
    if (slideImage) slideImage.style.width = `${zoom}%`;
  });

  document.addEventListener('keydown', (event) => {
    if (event.target.matches('input, textarea, select')) return;
    if (event.key === 'ArrowLeft') selectSlide(selectedSlide - 1);
    if (event.key === 'ArrowRight') selectSlide(selectedSlide + 1);
  });

  activateRightTab('design');
  updateBuilderControls();
  renderAll();
})();
