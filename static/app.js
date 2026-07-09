function safeLocalStorageGet(key) {
  try {
    return window.localStorage.getItem(key);
  } catch (_error) {
    return "";
  }
}

function safeLocalStorageSet(key, value) {
  try {
    window.localStorage.setItem(key, value);
  } catch (_error) {
    // Browser privacy settings can disable localStorage. The UI remains usable without it.
  }
}

function initializeStoreRequestApp() {
  const imageInput = document.querySelector("[data-image-input]");
  const imagePreview = document.querySelector("[data-image-preview]");
  const clearImagesButton = document.querySelector("[data-clear-images]");
  const fileInput = document.querySelector("[data-file-input]");
  const filePreview = document.querySelector("[data-file-preview]");
  const clearFilesButton = document.querySelector("[data-clear-files]");
  const extraImageInput = document.querySelector("[data-extra-image-input]");
  const extraImagePreview = document.querySelector("[data-extra-image-preview]");
  const clearExtraImagesButton = document.querySelector("[data-clear-extra-images]");
  const extraFileInput = document.querySelector("[data-extra-file-input]");
  const extraFilePreview = document.querySelector("[data-extra-file-preview]");
  const clearExtraFilesButton = document.querySelector("[data-clear-extra-files]");
  const copyTicketButtons = document.querySelectorAll("[data-copy-ticket]");
  const requestTypeSelect = document.querySelector("[data-request-type-select]");
  const requestTypeHint = document.querySelector("[data-request-type-hint]");
  const requestTypeTemplateCard = document.querySelector("[data-request-type-template-card]");
  const templateTitle = document.querySelector("[data-template-title]");
  const templateDescription = document.querySelector("[data-template-description]");
  const templateRecommended = document.querySelector("[data-template-recommended]");
  const templateRequired = document.querySelector("[data-template-required]");
  const templateHelp = document.querySelector("[data-template-help]");
  const descriptionRequiredMark = document.querySelector("[data-description-required]");
  const descriptionInput = document.querySelector("[data-description-input]");
  const descriptionTemplateHint = document.querySelector("[data-description-template-hint]");
  const expectedFinishInput = document.querySelector("[data-expected-finish-input]");
  const expectedFinishHint = document.querySelector("[data-expected-finish-hint]");
  const imageRequiredHint = document.querySelector("[data-image-required-hint]");
  const fileRequiredHint = document.querySelector("[data-file-required-hint]");
  const notificationRoot = document.querySelector("[data-notification-root]");
  const notificationToggle = document.querySelector("[data-notification-toggle]");
  const notificationPanel = document.querySelector("[data-notification-panel]");
  const notificationCount = document.querySelector("[data-notification-count]");
  const notificationList = document.querySelector("[data-notification-list]");
  const notificationReadAll = document.querySelector("[data-notification-read-all]");
  const notificationDesktopButton = document.querySelector("[data-notification-enable-desktop]");
  const notificationToasts = document.querySelector("[data-notification-toasts]");
  const modalOpenButtons = Array.from(document.querySelectorAll("[data-modal-open]"));
  const modalCloseButtons = Array.from(document.querySelectorAll("[data-modal-close]"));
  const modalOverlays = Array.from(document.querySelectorAll("[data-modal-overlay]"));
  const drawerOpenButtons = Array.from(document.querySelectorAll("[data-drawer-open]"));
  const drawerCloseButtons = Array.from(document.querySelectorAll("[data-drawer-close]"));
  const drawerOverlays = Array.from(document.querySelectorAll("[data-drawer]"));
  const employeeStoreForms = Array.from(document.querySelectorAll("[data-employee-store-form]"));
  const shiftScopeForms = Array.from(document.querySelectorAll("[data-shift-scope-form]"));
  const feedbackForms = Array.from(document.querySelectorAll("[data-feedback-form]"));
  const confirmForms = Array.from(
    new Set(
      Array.from(document.querySelectorAll("form[data-confirm-message]")).concat(
        Array.from(document.querySelectorAll("[data-confirm-message]"))
          .map((element) => element.closest("form"))
          .filter(Boolean),
      ),
    ),
  );
  const residueCleanupForms = Array.from(document.querySelectorAll("[data-residue-cleanup-form]"));
  const preserveScrollForms = Array.from(document.querySelectorAll("form[data-preserve-scroll]"));
  const cleanEmptyQueryForms = Array.from(document.querySelectorAll("form[data-clean-empty-query]"));
  const filterToggleButtons = Array.from(document.querySelectorAll("[data-filter-toggle]"));
  const restShiftButton = document.querySelector("[data-select-rest-shift]");
  const copyYesterdayButton = document.querySelector("[data-copy-yesterday-ui]");
  const adminTicketCreateForms = Array.from(document.querySelectorAll("[data-admin-ticket-create-form]"));
  const scheduleBulkForm = document.querySelector("[data-schedule-bulk-form]");
  const scheduleBulkSummary = document.querySelector("[data-schedule-bulk-summary]");
  const bulkDateCards = Array.from(document.querySelectorAll("[data-bulk-date-card]"));
  const selectAllEmployeesButton = document.querySelector("[data-select-all-employees]");
  const selectPrimaryEmployeesButton = document.querySelector("[data-select-primary-employees]");
  const selectSupportEmployeesButton = document.querySelector("[data-select-support-employees]");
  const clearEmployeesButton = document.querySelector("[data-clear-employees]");
  const selectAllScheduleDatesButton = document.querySelector("[data-select-all-schedule-dates]");
  const selectWeekdaysButton = document.querySelector("[data-select-weekdays]");
  const selectWeekendsButton = document.querySelector("[data-select-weekends]");
  const selectHolidaysButton = document.querySelector("[data-select-holidays]");
  const clearScheduleDatesButton = document.querySelector("[data-clear-schedule-dates]");
  const customScheduleModeInputs = Array.from(document.querySelectorAll("[data-custom-schedule-mode]"));
  const customScheduleFields = document.querySelector("[data-custom-schedule-fields]");
  const recalculateCustomDurationButton = document.querySelector("[data-recalculate-custom-duration]");
  const rolePermissionToolbar = document.querySelector("[data-role-permission-toolbar]");
  const rolePermissionSearch = document.querySelector("[data-role-permission-search]");
  const roleModuleFilterButtons = Array.from(document.querySelectorAll("[data-role-module-filter]"));
  const rolePermissionCards = Array.from(document.querySelectorAll("[data-role-permission-card]"));
  const roleListSearch = document.querySelector("[data-role-list-search]");
  const roleListFilterButtons = Array.from(document.querySelectorAll("[data-role-filter]"));
  const roleListItems = Array.from(document.querySelectorAll("[data-role-list-item]"));
  const roleEditorPanels = Array.from(document.querySelectorAll("[data-role-editor-panel]"));
  const dashboardFilterForm = document.querySelector("[data-dashboard-filter-form]");
  const advancedFilterPanel = document.querySelector("[data-advanced-filter-panel]");
  const advancedFilterToggle = document.querySelector("[data-advanced-filter-toggle]");
  const filterDropdowns = Array.from(document.querySelectorAll("[data-filter-dropdown]"));
  const dateRangeSelects = Array.from(document.querySelectorAll("[data-date-range-select]"));
  const customDateRanges = Array.from(document.querySelectorAll("[data-custom-date-range]"));

  let selectedImages = [];
  let selectedFiles = [];
  let selectedExtraImages = [];
  let selectedExtraFiles = [];
  let notificationLatestId = 0;
  let notificationInitialLoaded = false;
  let notificationDesktopEnabled = safeLocalStorageGet("storeRequestDesktopNotifications") === "1";
  let activeRolePermissionModule = "all";
  let activeRoleListFilter = "all";

  function formatFileSize(bytes) {
    if (bytes >= 1024 * 1024) {
      return `${(bytes / (1024 * 1024)).toFixed(1)}MB`;
    }
    if (bytes >= 1024) {
      return `${(bytes / 1024).toFixed(1)}KB`;
    }
    return `${bytes}B`;
  }

  function syncInputFiles(input, files) {
    if (!input || typeof DataTransfer === "undefined") {
      return;
    }
    const transfer = new DataTransfer();
    files.forEach((file) => transfer.items.add(file));
    input.files = transfer.files;
  }

  function createRemoveButton(onRemove) {
    const button = document.createElement("button");
    button.className = "ghost-button compact preview-remove";
    button.type = "button";
    button.textContent = "删除";
    button.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      onRemove();
    });
    return button;
  }

  function createFileMeta(file) {
    const meta = document.createElement("div");
    meta.className = "preview-meta";

    const name = document.createElement("span");
    name.className = "preview-name";
    name.textContent = file.name;
    meta.appendChild(name);

    const size = document.createElement("span");
    size.className = "preview-size";
    size.textContent = formatFileSize(file.size);
    meta.appendChild(size);

    return meta;
  }

  function setTemporaryButtonText(button, text) {
    const originalText = button.dataset.originalText || button.textContent;
    button.dataset.originalText = originalText;
    button.textContent = text;
    window.setTimeout(() => {
      button.textContent = originalText;
    }, 1600);
  }

  function findModalTarget(targetName) {
    if (!targetName) {
      return null;
    }
    const byId = document.getElementById(targetName);
    if (byId) {
      return byId;
    }
    return Array.from(document.querySelectorAll("[data-modal]")).find((modal) => modal.dataset.modal === targetName) || null;
  }

  function openModal(modal) {
    if (!modal) {
      return;
    }
    modal.hidden = false;
    modal.classList.add("is-open");
    const firstInput = modal.querySelector("input:not([type='hidden']), select, textarea, button");
    if (firstInput) {
      window.setTimeout(() => firstInput.focus(), 0);
    }
  }

  function closeModal(modal) {
    if (!modal) {
      return;
    }
    modal.classList.remove("is-open");
    modal.hidden = true;
  }

  window.openModal = (targetName) => openModal(findModalTarget(targetName));
  window.closeModal = (targetName) => closeModal(findModalTarget(targetName));

  function findDrawerTarget(targetName) {
    if (!targetName) {
      return null;
    }
    const byId = document.getElementById(targetName);
    if (byId) {
      return byId;
    }
    return Array.from(document.querySelectorAll("[data-drawer]")).find((drawer) => drawer.dataset.drawer === targetName) || null;
  }

  function openDrawer(drawer) {
    if (!drawer) {
      return;
    }
    drawer.hidden = false;
    drawer.classList.add("is-open");
    const firstInput = drawer.querySelector("input:not([type='hidden']), select, textarea, button");
    if (firstInput) {
      window.setTimeout(() => firstInput.focus(), 0);
    }
  }

  function closeDrawer(drawer) {
    if (!drawer) {
      return;
    }
    drawer.classList.remove("is-open");
    drawer.hidden = true;
  }

  function showAppToast(message, tone = "success") {
    if (!message) {
      return;
    }
    let toastArea = document.querySelector("[data-app-toast-area]");
    if (!toastArea) {
      toastArea = document.createElement("div");
      toastArea.className = "app-toast-area toast-container";
      toastArea.dataset.appToastArea = "1";
      document.body.appendChild(toastArea);
    }
    const toast = document.createElement("div");
    toast.className = `app-toast toast-card ${tone}`;
    toast.textContent = message;
    toastArea.appendChild(toast);
    window.setTimeout(() => toast.remove(), 1800);
  }

  function filterDropdownCheckboxes(dropdown) {
    return Array.from(dropdown.querySelectorAll('input[type="checkbox"]'));
  }

  function filterDropdownOptions(dropdown) {
    return Array.from(dropdown.querySelectorAll("[data-filter-option]"));
  }

  function checkedFilterDropdownLabels(dropdown) {
    return filterDropdownCheckboxes(dropdown)
      .filter((input) => input.checked)
      .map((input) => {
        const option = input.closest("[data-filter-option]");
        const label = option ? option.querySelector("span") : null;
        return label ? label.textContent.trim() : input.value;
      })
      .filter(Boolean);
  }

  function updateFilterDropdownSummary(dropdown) {
    if (!dropdown) {
      return;
    }
    const labels = checkedFilterDropdownLabels(dropdown);
    const labelTarget = dropdown.querySelector("[data-filter-dropdown-label]");
    const countTarget = dropdown.querySelector("[data-filter-selected-count]");
    const emptyLabel = dropdown.dataset.emptyLabel || "全部";
    const countLabel = dropdown.dataset.countLabel || "";
    if (labelTarget) {
      if (labels.length === 0) {
        labelTarget.textContent = emptyLabel;
      } else if (labels.length === 1) {
        labelTarget.textContent = labels[0];
      } else {
        labelTarget.textContent = `${labels[0]} +${labels.length - 1}`;
      }
      labelTarget.title = labels.length > 0 ? labels.join("、") : emptyLabel;
    }
    if (countTarget) {
      countTarget.textContent = String(labels.length);
    }
    dropdown.classList.toggle("has-selection", labels.length > 0);
    if (labels.length > 0 && countLabel) {
      dropdown.dataset.summary = `已选 ${labels.length} 个${countLabel}`;
    } else {
      dropdown.dataset.summary = emptyLabel;
    }
  }

  function snapshotFilterDropdown(dropdown) {
    const values = filterDropdownCheckboxes(dropdown)
      .filter((input) => input.checked)
      .map((input) => input.value);
    dropdown.dataset.filterSnapshot = JSON.stringify(values);
  }

  function restoreFilterDropdownSnapshot(dropdown) {
    let values = [];
    try {
      values = JSON.parse(dropdown.dataset.filterSnapshot || "[]");
    } catch (_error) {
      values = [];
    }
    const selected = new Set(values.map(String));
    filterDropdownCheckboxes(dropdown).forEach((input) => {
      input.checked = selected.has(String(input.value));
    });
    updateFilterDropdownSummary(dropdown);
  }

  function closeFilterDropdown(dropdown, restore = false) {
    if (!dropdown || !dropdown.open) {
      return;
    }
    if (restore) {
      restoreFilterDropdownSnapshot(dropdown);
    }
    const searchInput = dropdown.querySelector("[data-filter-dropdown-search]");
    if (searchInput) {
      searchInput.value = "";
    }
    filterDropdownOptions(dropdown).forEach((option) => option.classList.remove("is-hidden"));
    dropdown.open = false;
  }

  function closeOtherFilterDropdowns(currentDropdown) {
    filterDropdowns.forEach((dropdown) => {
      if (dropdown !== currentDropdown) {
        closeFilterDropdown(dropdown, false);
      }
    });
  }

  function setVisibleFilterDropdownOptionsChecked(dropdown, checked) {
    filterDropdownOptions(dropdown).forEach((option) => {
      if (option.classList.contains("is-hidden")) {
        return;
      }
      const input = option.querySelector('input[type="checkbox"]');
      if (input) {
        input.checked = checked;
      }
    });
    updateFilterDropdownSummary(dropdown);
  }

  function filterDropdownOptionList(dropdown, query) {
    const cleanQuery = String(query || "").trim().toLowerCase();
    filterDropdownOptions(dropdown).forEach((option) => {
      option.classList.toggle("is-hidden", Boolean(cleanQuery) && !option.textContent.toLowerCase().includes(cleanQuery));
    });
  }

  function updateAdvancedFilterToggle() {
    if (!advancedFilterToggle || !advancedFilterPanel) {
      return;
    }
    advancedFilterToggle.textContent = advancedFilterPanel.open ? "收起筛选" : "展开筛选";
  }

  function updateCustomDateRangeVisibility() {
    const isCustom = dateRangeSelects.some((select) => select.value === "custom");
    customDateRanges.forEach((range) => {
      range.hidden = !isCustom;
      Array.from(range.querySelectorAll("input")).forEach((input) => {
        input.disabled = !isCustom;
      });
    });
  }

  function shouldPreserveAdminScroll() {
    const path = window.location.pathname;
    return [
      "/admin",
      "/admin/schedules",
      "/admin/employees",
      "/admin/shift-types",
      "/admin/archive",
      "/admin/trash",
    ].includes(path);
  }

  function adminScrollStorageKey() {
    return `storeRequestScroll:${window.location.pathname}`;
  }

  function saveAdminScrollPosition() {
    if (!shouldPreserveAdminScroll()) {
      return;
    }
    try {
      window.sessionStorage.setItem(adminScrollStorageKey(), String(window.scrollY || 0));
    } catch (_error) {
      // Session storage can be disabled. Forms still submit normally.
    }
  }

  function restoreAdminScrollPosition() {
    if (!shouldPreserveAdminScroll()) {
      return;
    }
    if (window.location.hash) {
      const target = document.getElementById(window.location.hash.slice(1));
      if (target) {
        window.setTimeout(() => target.scrollIntoView({ block: "start" }), 0);
      }
      try {
        window.sessionStorage.removeItem(adminScrollStorageKey());
      } catch (_error) {
        // Ignore storage cleanup failures.
      }
      return;
    }
    let savedY = "";
    try {
      savedY = window.sessionStorage.getItem(adminScrollStorageKey()) || "";
      window.sessionStorage.removeItem(adminScrollStorageKey());
    } catch (_error) {
      savedY = "";
    }
    if (savedY) {
      const y = Number(savedY);
      if (!Number.isNaN(y)) {
        window.setTimeout(() => window.scrollTo(0, y), 0);
      }
    }
  }

  function setModalError(form, message) {
    const errorBox = form ? form.querySelector("[data-modal-error]") : null;
    if (!errorBox) {
      return;
    }
    errorBox.textContent = message || "";
    errorBox.hidden = !message;
  }

  function syncEmployeePrimaryStoreChoices(form) {
    const primarySelect = form.querySelector("[data-primary-store-select]");
    if (!primarySelect) {
      return;
    }
    const primaryStore = primarySelect.value || "";
    Array.from(form.querySelectorAll("[data-store-checkbox]")).forEach((checkbox) => {
      const chip = checkbox.closest("[data-store-choice]");
      const note = chip ? chip.querySelector("[data-primary-store-note]") : null;
      const isPrimaryStore = Boolean(primaryStore && checkbox.value === primaryStore);
      if (isPrimaryStore) {
        checkbox.checked = true;
        checkbox.dataset.primaryAutoChecked = "1";
      } else {
        if (checkbox.dataset.primaryAutoChecked === "1") {
          checkbox.checked = false;
          delete checkbox.dataset.primaryAutoChecked;
        }
      }
      if (chip) {
        chip.classList.toggle("is-primary-store-choice", isPrimaryStore);
      }
      if (note) {
        note.hidden = !isPrimaryStore;
      }
    });
  }

  function updateShiftScopeForm(form) {
    const checkedScope = form.querySelector('input[name="shift_scope"]:checked');
    const storeSelect = form.querySelector("[data-shift-store-select]");
    const isGlobalScope = checkedScope ? checkedScope.value === "global" : false;
    if (storeSelect) {
      storeSelect.disabled = isGlobalScope;
      storeSelect.required = !isGlobalScope;
      if (isGlobalScope) {
        storeSelect.value = "";
      }
    }
    form.classList.toggle("is-global-shift-scope", isGlobalScope);
  }

  function scheduleEmployeeInputs() {
    return scheduleBulkForm ? Array.from(scheduleBulkForm.querySelectorAll('input[name="employee_ids"]')) : [];
  }

  function scheduleDateInputs() {
    return scheduleBulkForm ? Array.from(scheduleBulkForm.querySelectorAll('input[name="schedule_dates"]')) : [];
  }

  function updateScheduleBulkSummary() {
    if (!scheduleBulkForm || !scheduleBulkSummary) {
      return;
    }
    const employeeCount = scheduleEmployeeInputs().filter((input) => input.checked).length;
    const dateCount = scheduleDateInputs().filter((input) => input.checked).length;
    const total = employeeCount * dateCount;
    const maxCount = Number(scheduleBulkForm.dataset.maxBulkScheduleCount || 0);
    scheduleBulkSummary.textContent = `已选 ${employeeCount} 名员工 × ${dateCount} 天 = 将生成 ${total} 条排班`;
    scheduleBulkSummary.classList.toggle("is-over-limit", Boolean(maxCount && total > maxCount));
  }

  function selectedScheduleMode() {
    const checkedMode = customScheduleModeInputs.find((input) => input.checked);
    return checkedMode ? checkedMode.value : "shift";
  }

  function customDurationInput() {
    return scheduleBulkForm ? scheduleBulkForm.querySelector('input[name="custom_duration_hours"]') : null;
  }

  function customTimeInput(name) {
    return scheduleBulkForm ? scheduleBulkForm.querySelector(`input[name="${name}"]`) : null;
  }

  function minutesFromTime(value) {
    const match = String(value || "").match(/^(\d{2}):(\d{2})$/);
    if (!match) {
      return null;
    }
    return Number(match[1]) * 60 + Number(match[2]);
  }

  function calculateCustomDurationHours() {
    const startInput = customTimeInput("custom_start_time");
    const endInput = customTimeInput("custom_end_time");
    const breakInput = customTimeInput("custom_break_minutes");
    const start = minutesFromTime(startInput ? startInput.value : "");
    let end = minutesFromTime(endInput ? endInput.value : "");
    if (start === null || end === null) {
      return "";
    }
    if (end <= start) {
      end += 24 * 60;
    }
    const breakMinutes = Math.max(Number(breakInput ? breakInput.value || 0 : 0), 0);
    const hours = Math.max((end - start - breakMinutes) / 60, 0);
    return hours > 0 ? String(Math.round(hours * 100) / 100) : "";
  }

  function recalculateCustomDuration(force = false) {
    const durationInput = customDurationInput();
    if (!durationInput) {
      return;
    }
    if (!force && durationInput.dataset.manualOverride === "1") {
      return;
    }
    const calculated = calculateCustomDurationHours();
    if (calculated) {
      durationInput.value = calculated;
      durationInput.dataset.autoCalculated = "1";
    }
  }

  function updateCustomScheduleFields() {
    if (!customScheduleModeInputs.length || !customScheduleFields) {
      return;
    }
    const isCustom = selectedScheduleMode() === "custom";
    customScheduleFields.hidden = !isCustom;
    Array.from(customScheduleFields.querySelectorAll("input, select, textarea")).forEach((field) => {
      field.disabled = !isCustom;
    });
    const shiftSelect = scheduleBulkForm ? scheduleBulkForm.querySelector("[data-shift-select]") : null;
    if (shiftSelect) {
      shiftSelect.disabled = isCustom;
      shiftSelect.required = !isCustom;
    }
    if (isCustom) {
      recalculateCustomDuration(false);
    }
  }

  function setScheduleInputsChecked(inputs, checked) {
    inputs.forEach((input) => {
      input.checked = checked;
      syncScheduleDateCard(input);
    });
    updateScheduleBulkSummary();
  }

  function syncScheduleDateCard(input) {
    const card = input ? input.closest("[data-bulk-date-card], .schedule-date-chip") : null;
    if (card) {
      card.classList.toggle("is-selected", Boolean(input.checked));
    }
  }

  function setScheduleEmployeesByRole(role) {
    scheduleEmployeeInputs().forEach((input) => {
      const chip = input.closest(".schedule-employee-chip");
      if (!chip) {
        input.checked = false;
        return;
      }
      input.checked = role === "primary" ? chip.hasAttribute("data-primary-store-employee") : chip.hasAttribute("data-support-store-employee");
    });
    updateScheduleBulkSummary();
  }

  function setScheduleDatesByWeekend(weekendOnly) {
    scheduleDateInputs().forEach((input) => {
      const chip = input.closest("[data-weekend]");
      const isWeekend = chip ? chip.dataset.weekend === "1" : false;
      input.checked = weekendOnly ? isWeekend : !isWeekend;
      syncScheduleDateCard(input);
    });
    updateScheduleBulkSummary();
  }

  function setScheduleDatesByHoliday() {
    scheduleDateInputs().forEach((input) => {
      const chip = input.closest("[data-holiday]");
      input.checked = chip ? chip.dataset.holiday === "1" : false;
      syncScheduleDateCard(input);
    });
    updateScheduleBulkSummary();
  }

  function rolePermissionMatchesSearch(item, query) {
    if (!query) {
      return true;
    }
    return String(item.dataset.permissionSearch || "").toLowerCase().includes(query);
  }

  function updateRolePermissionMatrix() {
    if (!rolePermissionCards.length) {
      return;
    }
    const query = rolePermissionSearch ? rolePermissionSearch.value.trim().toLowerCase() : "";
    rolePermissionCards.forEach((card) => {
      card.querySelectorAll("[data-permission-item]").forEach((item) => {
        const moduleName = item.dataset.permissionModule || "";
        const moduleMatches = activeRolePermissionModule === "all" || moduleName === activeRolePermissionModule;
        item.hidden = !(moduleMatches && rolePermissionMatchesSearch(item, query));
      });
      card.querySelectorAll("[data-permission-module-group]").forEach((group) => {
        const moduleName = group.dataset.permissionModule || group.dataset.permissionModuleGroup || "";
        const visibleItems = Array.from(group.querySelectorAll("[data-permission-item]")).filter((item) => !item.hidden);
        const moduleVisible = visibleItems.length > 0;
        group.hidden = false;
        group.dataset.filterHidden = moduleVisible ? "false" : "true";
        group.classList.toggle("is-filtered-out", !moduleVisible);
        group.classList.toggle("is-filter-hidden", !moduleVisible);
        group.dataset.moduleFilterCurrent = activeRolePermissionModule;
      });
      const page = card.closest("[data-role-permission-page]");
      if (page) {
        page.dataset.moduleFilterCurrent = activeRolePermissionModule;
      }
    });
  }

  function checkedRolePermissionValues(form) {
    if (!form) {
      return [];
    }
    return Array.from(form.querySelectorAll('input[name="permissions"]:checked'))
      .map((input) => input.value)
      .sort();
  }

  function serializedRolePermissions(form) {
    return checkedRolePermissionValues(form).join("|");
  }

  function panelForRoleForm(form) {
    return form ? form.closest("[data-role-editor-panel]") : null;
  }

  function listItemForRolePanel(panel) {
    if (!panel) {
      return null;
    }
    const panelId = panel.dataset.rolePanelId || panel.id || "";
    return roleListItems.find((item) => item.dataset.roleTarget === panelId) || null;
  }

  function activeRoleEditorPanel() {
    return roleEditorPanels.find((panel) => !panel.hidden) || roleEditorPanels[0] || null;
  }

  function updateRolePermissionCounts(form) {
    const selectedCount = checkedRolePermissionValues(form).length;
    const panel = panelForRoleForm(form);
    if (panel) {
      const selectedNode = panel.querySelector("[data-role-selected-count]");
      if (selectedNode) {
        selectedNode.textContent = String(selectedCount);
      }
      const listItem = listItemForRolePanel(panel);
      const listCount = listItem ? listItem.querySelector("[data-role-list-count]") : null;
      if (listCount) {
        listCount.textContent = `${selectedCount} 项权限`;
      }
    }
    form.querySelectorAll("[data-permission-module-group]").forEach((group) => {
      const moduleSelected = group.querySelectorAll('input[name="permissions"]:checked').length;
      const moduleCount = group.querySelector("[data-module-selected-count]");
      if (moduleCount) {
        moduleCount.textContent = String(moduleSelected);
      }
    });
  }

  function syncRoleFormDirtyState(form) {
    if (!form) {
      return false;
    }
    const dirty = serializedRolePermissions(form) !== (form.dataset.roleInitialPermissions || "");
    form.dataset.roleDirty = dirty ? "1" : "0";
    const panel = panelForRoleForm(form);
    if (panel) {
      panel.classList.toggle("is-dirty", dirty);
      const listItem = listItemForRolePanel(panel);
      if (listItem) {
        listItem.classList.toggle("is-dirty", dirty);
      }
    }
    return dirty;
  }

  function updateRoleFormState(form) {
    if (!form) {
      return;
    }
    updateRolePermissionCounts(form);
    syncRoleFormDirtyState(form);
  }

  function initializeRoleForm(form) {
    if (!form) {
      return;
    }
    form.dataset.roleInitialPermissions = serializedRolePermissions(form);
    form.dataset.roleDirty = "0";
    updateRolePermissionCounts(form);
  }

  function roleListItemMatchesFilter(item, query) {
    const roleName = String(item.dataset.roleName || item.textContent || "").toLowerCase();
    const category = item.dataset.roleCategory || "";
    const isSystem = item.dataset.roleSystem === "1";
    const queryMatches = !query || roleName.includes(query);
    let filterMatches = true;
    if (activeRoleListFilter === "current") {
      filterMatches = category === "current";
    } else if (activeRoleListFilter === "legacy") {
      filterMatches = category === "legacy";
    } else if (activeRoleListFilter === "system") {
      filterMatches = isSystem;
    }
    return queryMatches && filterMatches;
  }

  function updateRoleList() {
    if (!roleListItems.length) {
      return;
    }
    const query = roleListSearch ? roleListSearch.value.trim().toLowerCase() : "";
    roleListItems.forEach((item) => {
      item.hidden = !roleListItemMatchesFilter(item, query);
    });
  }

  function hasDirtyRoleForm() {
    return roleEditorPanels.some((panel) => {
      const form = panel.querySelector("[data-role-form]");
      return form && form.matches('[data-role-dirty="1"]');
    });
  }

  function switchRolePanel(targetPanelId) {
    const targetPanel = roleEditorPanels.find((panel) => (panel.dataset.rolePanelId || panel.id) === targetPanelId);
    if (!targetPanel) {
      return;
    }
    const currentPanel = activeRoleEditorPanel();
    if (currentPanel && currentPanel !== targetPanel) {
      const currentForm = currentPanel.querySelector("[data-role-form]");
      if (
        currentForm &&
        currentForm.dataset.roleDirty === "1" &&
        !window.confirm("当前角色有未保存修改，切换后本页会暂存勾选，但离开或刷新前仍需保存。是否继续切换？")
      ) {
        return;
      }
    }
    roleEditorPanels.forEach((panel) => {
      panel.hidden = panel !== targetPanel;
    });
    roleListItems.forEach((item) => {
      item.classList.toggle("active", item.dataset.roleTarget === targetPanelId);
    });
    updateRolePermissionMatrix();
  }

  function escapeCssValue(value) {
    if (window.CSS && typeof window.CSS.escape === "function") {
      return window.CSS.escape(value);
    }
    return String(value).replace(/["\\]/g, "\\$&");
  }

  function syncRolePermissionCheckboxes(sourceInput) {
    if (!sourceInput || !sourceInput.form) {
      return;
    }
    const selector = `input[name="permissions"][value="${escapeCssValue(sourceInput.value)}"]`;
    sourceInput.form.querySelectorAll(selector).forEach((input) => {
      input.checked = sourceInput.checked;
    });
    updateRoleFormState(sourceInput.form);
  }

  function currentRoleModuleItems(form) {
    if (!form) {
      return [];
    }
    return Array.from(form.querySelectorAll("[data-permission-item]")).filter((item) => {
      const moduleName = item.dataset.permissionModule || "";
      return activeRolePermissionModule === "all" || moduleName === activeRolePermissionModule;
    });
  }

  function setCurrentRoleModuleChecked(form, checked) {
    const changedValues = new Set();
    currentRoleModuleItems(form).forEach((item) => {
      const input = item.querySelector('input[name="permissions"]');
      if (input) {
        input.checked = checked;
        changedValues.add(input.value);
      }
    });
    changedValues.forEach((value) => {
      form.querySelectorAll(`input[name="permissions"][value="${escapeCssValue(value)}"]`).forEach((input) => {
        input.checked = checked;
      });
    });
    updateRoleFormState(form);
  }

  function updateRequestTypeHint() {
    if (!requestTypeSelect || !requestTypeHint) {
      return;
    }
    let rules = {};
    try {
      rules = JSON.parse(requestTypeSelect.dataset.requestTypeRules || "{}");
    } catch (_error) {
      rules = {};
    }
    const selectedRule = rules[requestTypeSelect.value] || {};
    const fallbackRecommended = ["品牌", "商品名称", "规格条码", "数量"];
    const recommendedFields = Array.isArray(selectedRule.recommended_fields) && selectedRule.recommended_fields.length
      ? selectedRule.recommended_fields
      : fallbackRecommended;
    const requiredFields = Array.isArray(selectedRule.required_fields) && selectedRule.required_fields.length
      ? selectedRule.required_fields
      : ["问题说明"];
    const helpText = selectedRule.field_help_text || selectedRule.description_hint || "商品信息是建议填写，不完整也可以先提交。";
    requestTypeHint.textContent = helpText;

    if (requestTypeTemplateCard) {
      requestTypeTemplateCard.hidden = false;
    }
    if (templateTitle) {
      templateTitle.textContent = selectedRule.template_name || requestTypeSelect.value || "请选择需求类型";
    }
    if (templateDescription) {
      templateDescription.textContent = selectedRule.description || "系统会根据需求类型提示建议填写信息和必填信息。";
    }
    const renderTags = (container, items, emptyText) => {
      if (!container) {
        return;
      }
      container.innerHTML = "";
      const values = items.length ? items : [emptyText];
      values.forEach((item) => {
        const tag = document.createElement("span");
        tag.className = "tag-chip";
        tag.textContent = item;
        container.appendChild(tag);
      });
    };
    renderTags(templateRecommended, recommendedFields, "无");
    renderTags(templateRequired, requiredFields, "无");
    if (templateHelp) {
      templateHelp.textContent = helpText;
    }

    const descriptionRequired = selectedRule.required_description !== 0;
    if (descriptionRequiredMark) {
      descriptionRequiredMark.hidden = !descriptionRequired;
    }
    if (descriptionInput) {
      descriptionInput.required = descriptionRequired;
    }
    if (descriptionTemplateHint) {
      descriptionTemplateHint.textContent = descriptionRequired ? "此类型需要填写问题说明。" : "此类型的问题说明可选，但建议写清背景。";
    }
    const expectedRequired = selectedRule.expected_finish_required === 1;
    if (expectedFinishInput) {
      expectedFinishInput.required = expectedRequired;
    }
    if (expectedFinishHint) {
      expectedFinishHint.textContent = expectedRequired ? "此类型必须填写期望完成时间。" : "如有明确期限可填写；未填写时可能按 SLA 自动计算。";
    }
    if (imageRequiredHint) {
      imageRequiredHint.hidden = selectedRule.image_required !== 1;
    }
    if (fileRequiredHint) {
      fileRequiredHint.hidden = selectedRule.file_required !== 1;
    }
  }

  function notificationSeverityLabel(severity) {
    if (severity === "urgent") {
      return "紧急";
    }
    if (severity === "warning") {
      return "提醒";
    }
    return "消息";
  }

  function notificationPost(url) {
    const formData = new FormData();
    formData.append("csrf_token", notificationRoot ? notificationRoot.dataset.csrfToken || "" : "");
    return fetch(url, {
      method: "POST",
      body: formData,
      credentials: "same-origin",
    });
  }

  async function fetchNotificationPayload(params = {}) {
    const query = new URLSearchParams();
    Object.entries(params).forEach(([key, value]) => {
      if (value !== undefined && value !== null && value !== "") {
        query.set(key, String(value));
      }
    });
    const response = await fetch(`/admin/api/notifications${query.toString() ? `?${query}` : ""}`, {
      credentials: "same-origin",
    });
    if (!response.ok) {
      throw new Error("notification fetch failed");
    }
    return response.json();
  }

  function updateNotificationCount(count) {
    if (!notificationCount) {
      return;
    }
    const value = Number(count || 0);
    notificationCount.textContent = value > 99 ? "99+" : String(value);
    notificationCount.hidden = value <= 0;
  }

  function notificationActionList(notification) {
    if (Array.isArray(notification.actions) && notification.actions.length > 0) {
      return notification.actions;
    }
    const actions = [];
    if (notification.detail_url) {
      actions.push({ label: "查看工单", url: notification.detail_url, method: "get" });
    }
    actions.push({
      label: "标记已读",
      url: `/admin/api/notifications/${notification.id}/read`,
      method: "post",
      disabled: Boolean(notification.is_read),
    });
    return actions;
  }

  function createNotificationActionControl(action, onComplete) {
    const method = String(action.method || "get").toLowerCase();
    if (method === "post") {
      const button = document.createElement("button");
      button.className = action.danger ? "danger-button compact" : "ghost-button compact";
      button.type = "button";
      button.textContent = action.label || "操作";
      button.disabled = Boolean(action.disabled);
      button.dataset.notificationActionUrl = action.url || "";
      if (typeof onComplete === "function") {
        button.addEventListener("click", async (event) => {
          event.preventDefault();
          if (!button.dataset.notificationActionUrl) {
            return;
          }
          await notificationPost(button.dataset.notificationActionUrl);
          await onComplete(action);
        });
      }
      return button;
    }
    const link = document.createElement("a");
    link.className = action.danger ? "danger-button compact" : "ghost-button compact";
    link.href = action.url || "#";
    link.textContent = action.label || "查看";
    return link;
  }

  function appendNotificationActions(container, notification, onComplete) {
    notificationActionList(notification).forEach((action) => {
      if (!action || !action.url) {
        return;
      }
      container.appendChild(createNotificationActionControl(action, onComplete));
    });
  }

  function createNotificationItem(notification) {
    const item = document.createElement("article");
    item.className = `notification-item ${notification.is_read ? "read" : "unread"}`;
    item.dataset.notificationId = String(notification.id);

    const top = document.createElement("div");
    top.className = "notification-item-top";

    const title = document.createElement("div");
    title.className = "notification-item-title";
    title.textContent = notification.title || "消息";
    top.appendChild(title);

    const severity = document.createElement("span");
    severity.className = `notification-severity ${notification.severity || "info"}`;
    severity.textContent = notificationSeverityLabel(notification.severity || "info");
    top.appendChild(severity);
    item.appendChild(top);

    const content = document.createElement("p");
    content.className = "notification-item-content";
    content.textContent = notification.content || "";
    item.appendChild(content);

    const meta = document.createElement("p");
    meta.className = "notification-item-meta";
    meta.textContent = `${notification.store_name || "-"} · ${notification.ticket_no || "-"} · ${notification.created_at || ""}`;
    item.appendChild(meta);

    const actions = document.createElement("div");
    actions.className = "notification-item-actions";
    appendNotificationActions(actions, notification);
    item.appendChild(actions);
    return item;
  }

  function renderNotificationList(notifications) {
    if (!notificationList) {
      return;
    }
    notificationList.innerHTML = "";
    if (!notifications || notifications.length === 0) {
      const empty = document.createElement("p");
      empty.className = "empty-note";
      empty.textContent = "暂无消息。";
      notificationList.appendChild(empty);
      return;
    }
    notifications.forEach((notification) => {
      notificationList.appendChild(createNotificationItem(notification));
    });
  }

  function scheduleToastRemoval(toast) {
    const timer = window.setTimeout(() => {
      if (toast.dataset.hovering !== "1") {
        toast.remove();
      }
    }, 8000);
    toast.dataset.timer = String(timer);
  }

  function showBrowserNotification(notification) {
    if (!notificationDesktopEnabled || !("Notification" in window) || Notification.permission !== "granted") {
      return;
    }
    const desktopNotification = new Notification(notification.title || "消息提醒", {
      body: notification.content || "",
      tag: `store-request-${notification.id}`,
    });
    desktopNotification.onclick = () => {
      window.focus();
      if (notification.detail_url) {
        window.location.href = notification.detail_url;
      }
    };
  }

  function showNotificationToast(notification) {
    if (!notificationToasts) {
      return;
    }
    while (notificationToasts.children.length >= 3) {
      notificationToasts.firstElementChild.remove();
    }
    const toast = document.createElement("article");
    toast.className = `notification-toast ${notification.severity || "info"}`;

    const top = document.createElement("div");
    top.className = "notification-toast-top";
    const title = document.createElement("div");
    title.className = "notification-toast-title";
    title.textContent = notification.title || "消息提醒";
    top.appendChild(title);
    const close = document.createElement("button");
    close.className = "ghost-button compact notification-toast-close";
    close.type = "button";
    close.textContent = "关闭";
    close.addEventListener("click", () => toast.remove());
    top.appendChild(close);
    toast.appendChild(top);

    const content = document.createElement("p");
    content.className = "notification-item-content";
    content.textContent = notification.content || "";
    toast.appendChild(content);

    const meta = document.createElement("p");
    meta.className = "notification-item-meta";
    meta.textContent = `${notification.store_name || "-"} · ${notification.ticket_no || "-"} · ${notification.created_at || ""}`;
    toast.appendChild(meta);

    const actions = document.createElement("div");
    actions.className = "notification-item-actions";
    appendNotificationActions(actions, notification, async () => {
      toast.remove();
      refreshNotifications();
    });
    toast.appendChild(actions);

    toast.addEventListener("mouseenter", () => {
      toast.dataset.hovering = "1";
      window.clearTimeout(Number(toast.dataset.timer || 0));
    });
    toast.addEventListener("mouseleave", () => {
      toast.dataset.hovering = "0";
      scheduleToastRemoval(toast);
    });
    notificationToasts.appendChild(toast);
    scheduleToastRemoval(toast);
    showBrowserNotification(notification);
  }

  async function refreshNotifications() {
    if (!notificationRoot) {
      return;
    }
    try {
      const payload = await fetchNotificationPayload({ limit: 20 });
      updateNotificationCount(payload.unread_count);
      renderNotificationList(payload.notifications || []);
      notificationLatestId = Math.max(notificationLatestId, Number(payload.latest_id || 0));
      notificationInitialLoaded = true;
    } catch (_error) {
      // Keep the admin page usable if the lightweight notification endpoint is unavailable.
    }
  }

  async function pollNotifications() {
    if (!notificationRoot || !notificationInitialLoaded) {
      return;
    }
    try {
      const payload = await fetchNotificationPayload({ after_id: notificationLatestId, limit: 20 });
      updateNotificationCount(payload.unread_count);
      const notifications = payload.notifications || [];
      if (notifications.length > 0) {
        notifications.slice().reverse().forEach(showNotificationToast);
        notificationLatestId = Math.max(notificationLatestId, Number(payload.latest_id || 0));
        await refreshNotifications();
      } else {
        notificationLatestId = Math.max(notificationLatestId, Number(payload.latest_id || 0));
      }
    } catch (_error) {
      // Polling is best-effort and should not disturb normal admin work.
    }
  }

  async function copyText(text) {
    if (navigator.clipboard && window.isSecureContext) {
      await navigator.clipboard.writeText(text);
      return true;
    }

    const textarea = document.createElement("textarea");
    textarea.value = text;
    textarea.setAttribute("readonly", "");
    textarea.style.position = "fixed";
    textarea.style.left = "-9999px";
    textarea.style.top = "0";
    document.body.appendChild(textarea);
    textarea.focus();
    textarea.select();
    try {
      return document.execCommand("copy");
    } finally {
      document.body.removeChild(textarea);
    }
  }

  const bulkRoot = document.querySelector("[data-bulk-root]");
  const ticketCheckboxes = Array.from(document.querySelectorAll("[data-ticket-select], [data-ticket-checkbox]"));
  const selectCurrentControls = Array.from(
    document.querySelectorAll("[data-select-current-page], [data-select-current-page-button]"),
  );
  const clearSelectionButton = document.querySelector("[data-clear-selection]");
  const selectFilteredButton = document.querySelector("[data-select-filtered]");
  const selectedCountElement = document.querySelector("[data-selected-count]");
  const filteredNotice = document.querySelector("[data-filtered-notice]");
  const bulkForms = Array.from(document.querySelectorAll("[data-bulk-form]"));
  let filteredSelectionActive = false;

  function selectedTicketIds() {
    return ticketCheckboxes.filter((checkbox) => checkbox.checked).map((checkbox) => checkbox.value);
  }

  function setBulkScope(scope) {
    document.querySelectorAll("[data-select-scope]").forEach((input) => {
      input.value = scope;
    });
  }

  function updateBulkSelectionState() {
    const selectedCount = selectedTicketIds().length;
    if (selectedCountElement) {
      selectedCountElement.textContent = String(selectedCount);
    }
    if (filteredNotice) {
      filteredNotice.hidden = !filteredSelectionActive;
    }
    const allChecked = ticketCheckboxes.length > 0 && selectedCount === ticketCheckboxes.length;
    selectCurrentControls.forEach((control) => {
      if (control.matches('input[type="checkbox"]')) {
        control.checked = allChecked;
        control.indeterminate = selectedCount > 0 && !allChecked;
      }
    });
  }

  if (bulkRoot && ticketCheckboxes.length > 0) {
    const filteredCount = Number(bulkRoot.dataset.filteredCount || "0");
    ticketCheckboxes.forEach((checkbox) => {
      checkbox.addEventListener("change", () => {
        filteredSelectionActive = false;
        setBulkScope("selected");
        updateBulkSelectionState();
      });
    });

    selectCurrentControls.forEach((control) => {
      if (control.matches('input[type="checkbox"]')) {
        control.addEventListener("change", () => {
          filteredSelectionActive = false;
          setBulkScope("selected");
          ticketCheckboxes.forEach((checkbox) => {
            checkbox.checked = control.checked;
          });
          updateBulkSelectionState();
        });
        return;
      }
      control.addEventListener("click", (event) => {
        event.preventDefault();
        filteredSelectionActive = false;
        setBulkScope("selected");
        ticketCheckboxes.forEach((checkbox) => {
          checkbox.checked = true;
        });
        updateBulkSelectionState();
      });
    });

    if (clearSelectionButton) {
      clearSelectionButton.addEventListener("click", (event) => {
        event.preventDefault();
        filteredSelectionActive = false;
        setBulkScope("selected");
        ticketCheckboxes.forEach((checkbox) => {
          checkbox.checked = false;
        });
        updateBulkSelectionState();
      });
    }

    if (selectFilteredButton) {
      selectFilteredButton.addEventListener("click", (event) => {
        event.preventDefault();
        if (filteredCount <= 0) {
          window.alert("当前筛选条件下没有可操作工单。");
          return;
        }
        filteredSelectionActive = true;
        setBulkScope("filtered");
        updateBulkSelectionState();
      });
    }

    bulkForms.forEach((form) => {
      form.addEventListener("submit", (event) => {
        const scope = filteredSelectionActive ? "filtered" : "selected";
        setBulkScope(scope);
        const ids = selectedTicketIds();
        if (scope === "selected") {
          if (ids.length === 0) {
            event.preventDefault();
            window.alert("请先选择工单。");
            return;
          }
        } else if (filteredCount <= 0) {
          event.preventDefault();
          window.alert("当前筛选条件下没有可操作工单。");
          return;
        }
        const submitter = event.submitter || document.activeElement;
        const message =
          scope === "filtered" ? submitter?.dataset?.confirmFiltered : submitter?.dataset?.confirmSelected;
        if (message && !window.confirm(message)) {
          event.preventDefault();
        }
      });
    });

    updateBulkSelectionState();
  }

  function renderImages() {
    if (!imagePreview) {
      return;
    }
    imagePreview.innerHTML = "";
    selectedImages.forEach((file, index) => {
      const item = document.createElement("div");
      item.className = "preview-item preview-item-image";

      const thumbnail = document.createElement("img");
      thumbnail.alt = file.name;
      thumbnail.src = URL.createObjectURL(file);
      thumbnail.onload = () => URL.revokeObjectURL(thumbnail.src);
      item.appendChild(thumbnail);

      item.appendChild(createFileMeta(file));
      item.appendChild(
        createRemoveButton(() => {
          selectedImages.splice(index, 1);
          renderImages();
          syncInputFiles(imageInput, selectedImages);
        }),
      );
      imagePreview.appendChild(item);
    });
  }

  function renderFiles() {
    if (!filePreview) {
      return;
    }
    filePreview.innerHTML = "";
    selectedFiles.forEach((file, index) => {
      const item = document.createElement("div");
      item.className = "preview-item preview-item-file";
      item.appendChild(createFileMeta(file));
      item.appendChild(
        createRemoveButton(() => {
          selectedFiles.splice(index, 1);
          renderFiles();
          syncInputFiles(fileInput, selectedFiles);
        }),
      );
      filePreview.appendChild(item);
    });
  }

  function renderExtraImages() {
    if (!extraImagePreview) {
      return;
    }
    extraImagePreview.innerHTML = "";
    selectedExtraImages.forEach((file, index) => {
      const item = document.createElement("div");
      item.className = "preview-item preview-item-image";

      const thumbnail = document.createElement("img");
      thumbnail.alt = file.name;
      thumbnail.src = URL.createObjectURL(file);
      thumbnail.onload = () => URL.revokeObjectURL(thumbnail.src);
      item.appendChild(thumbnail);

      item.appendChild(createFileMeta(file));
      item.appendChild(
        createRemoveButton(() => {
          selectedExtraImages.splice(index, 1);
          renderExtraImages();
          syncInputFiles(extraImageInput, selectedExtraImages);
        }),
      );
      extraImagePreview.appendChild(item);
    });
  }

  function renderExtraFiles() {
    if (!extraFilePreview) {
      return;
    }
    extraFilePreview.innerHTML = "";
    selectedExtraFiles.forEach((file, index) => {
      const item = document.createElement("div");
      item.className = "preview-item preview-item-file";
      item.appendChild(createFileMeta(file));
      item.appendChild(
        createRemoveButton(() => {
          selectedExtraFiles.splice(index, 1);
          renderExtraFiles();
          syncInputFiles(extraFileInput, selectedExtraFiles);
        }),
      );
      extraFilePreview.appendChild(item);
    });
  }

  if (imageInput && imagePreview) {
    imageInput.addEventListener("change", () => {
      selectedImages = selectedImages.concat(Array.from(imageInput.files || []));
      renderImages();
      syncInputFiles(imageInput, selectedImages);
    });
  }

  if (fileInput && filePreview) {
    fileInput.addEventListener("change", () => {
      selectedFiles = selectedFiles.concat(Array.from(fileInput.files || []));
      renderFiles();
      syncInputFiles(fileInput, selectedFiles);
    });
  }

  if (extraImageInput && extraImagePreview) {
    extraImageInput.addEventListener("change", () => {
      selectedExtraImages = selectedExtraImages.concat(Array.from(extraImageInput.files || []));
      renderExtraImages();
      syncInputFiles(extraImageInput, selectedExtraImages);
    });
  }

  if (extraFileInput && extraFilePreview) {
    extraFileInput.addEventListener("change", () => {
      selectedExtraFiles = selectedExtraFiles.concat(Array.from(extraFileInput.files || []));
      renderExtraFiles();
      syncInputFiles(extraFileInput, selectedExtraFiles);
    });
  }

  if (clearImagesButton && imageInput && imagePreview) {
    clearImagesButton.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      selectedImages = [];
      imageInput.value = "";
      syncInputFiles(imageInput, selectedImages);
      imagePreview.innerHTML = "";
    });
  }

  if (clearFilesButton && fileInput && filePreview) {
    clearFilesButton.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      selectedFiles = [];
      fileInput.value = "";
      syncInputFiles(fileInput, selectedFiles);
      filePreview.innerHTML = "";
    });
  }

  if (clearExtraImagesButton && extraImageInput && extraImagePreview) {
    clearExtraImagesButton.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      selectedExtraImages = [];
      extraImageInput.value = "";
      syncInputFiles(extraImageInput, selectedExtraImages);
      extraImagePreview.innerHTML = "";
    });
  }

  if (clearExtraFilesButton && extraFileInput && extraFilePreview) {
    clearExtraFilesButton.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      selectedExtraFiles = [];
      extraFileInput.value = "";
      syncInputFiles(extraFileInput, selectedExtraFiles);
      extraFilePreview.innerHTML = "";
    });
  }

  if (requestTypeSelect && requestTypeHint) {
    requestTypeSelect.addEventListener("change", updateRequestTypeHint);
    updateRequestTypeHint();
  }

  if (scheduleBulkForm) {
    updateCustomScheduleFields();
    customScheduleModeInputs.forEach((input) => {
      input.addEventListener("change", updateCustomScheduleFields);
    });
    ["custom_start_time", "custom_end_time", "custom_break_minutes"].forEach((name) => {
      const input = customTimeInput(name);
      if (input) {
        input.addEventListener("input", () => recalculateCustomDuration(false));
        input.addEventListener("change", () => recalculateCustomDuration(false));
      }
    });
    const durationInput = customDurationInput();
    if (durationInput) {
      durationInput.addEventListener("input", () => {
        durationInput.dataset.manualOverride = durationInput.value ? "1" : "0";
      });
    }
    if (recalculateCustomDurationButton) {
      recalculateCustomDurationButton.addEventListener("click", (event) => {
        event.preventDefault();
        const input = customDurationInput();
        if (input) {
          input.dataset.manualOverride = "0";
        }
        recalculateCustomDuration(true);
      });
    }
    scheduleEmployeeInputs().forEach((input) => input.addEventListener("change", updateScheduleBulkSummary));
    scheduleDateInputs().forEach((input) => {
      input.addEventListener("change", () => {
        syncScheduleDateCard(input);
        updateScheduleBulkSummary();
      });
      syncScheduleDateCard(input);
    });
    bulkDateCards.forEach((card) => {
      card.addEventListener("click", (event) => {
        const input = card.querySelector('input[name="schedule_dates"]');
        if (!input || input.disabled) {
          return;
        }
        event.preventDefault();
        input.checked = !input.checked;
        syncScheduleDateCard(input);
        updateScheduleBulkSummary();
      });
    });
  if (selectAllEmployeesButton) {
    selectAllEmployeesButton.addEventListener("click", () => setScheduleInputsChecked(scheduleEmployeeInputs(), true));
  }
  if (selectPrimaryEmployeesButton) {
    selectPrimaryEmployeesButton.addEventListener("click", () => setScheduleEmployeesByRole("primary"));
  }
  if (selectSupportEmployeesButton) {
    selectSupportEmployeesButton.addEventListener("click", () => setScheduleEmployeesByRole("support"));
  }
    if (clearEmployeesButton) {
      clearEmployeesButton.addEventListener("click", () => setScheduleInputsChecked(scheduleEmployeeInputs(), false));
    }
    if (selectAllScheduleDatesButton) {
      selectAllScheduleDatesButton.addEventListener("click", () => setScheduleInputsChecked(scheduleDateInputs(), true));
    }
    if (selectWeekdaysButton) {
      selectWeekdaysButton.addEventListener("click", () => setScheduleDatesByWeekend(false));
    }
    if (selectWeekendsButton) {
      selectWeekendsButton.addEventListener("click", () => setScheduleDatesByWeekend(true));
    }
    if (selectHolidaysButton) {
      selectHolidaysButton.addEventListener("click", () => setScheduleDatesByHoliday());
    }
    if (clearScheduleDatesButton) {
      clearScheduleDatesButton.addEventListener("click", () => setScheduleInputsChecked(scheduleDateInputs(), false));
    }
    scheduleBulkForm.addEventListener("submit", (event) => {
      const employeeCount = scheduleEmployeeInputs().filter((input) => input.checked).length;
      const dateCount = scheduleDateInputs().filter((input) => input.checked).length;
      const total = employeeCount * dateCount;
      const maxCount = Number(scheduleBulkForm.dataset.maxBulkScheduleCount || 0);
      if (maxCount && total > maxCount) {
        event.preventDefault();
        window.alert(`一次最多批量生成 ${maxCount} 条排班，请减少员工或日期数量。`);
      }
    });
    updateScheduleBulkSummary();
  }

  const copyLastWeekButton = document.querySelector("[data-copy-last-week-fill]");
  if (copyLastWeekButton) {
    copyLastWeekButton.addEventListener("click", (event) => {
      event.preventDefault();
      const form = copyLastWeekButton.closest("form");
      if (!form) {
        return;
      }
      const today = new Date();
      const day = today.getDay() || 7;
      const monday = new Date(today);
      monday.setDate(today.getDate() - day + 1);
      const sourceStart = new Date(monday);
      sourceStart.setDate(monday.getDate() - 7);
      const sourceEnd = new Date(monday);
      sourceEnd.setDate(monday.getDate() - 1);
      const toDateValue = (value) => value.toISOString().slice(0, 10);
      const sourceStartInput = form.querySelector('input[name="source_start_date"]');
      const sourceEndInput = form.querySelector('input[name="source_end_date"]');
      const targetStartInput = form.querySelector('input[name="target_start_date"]');
      if (sourceStartInput) sourceStartInput.value = toDateValue(sourceStart);
      if (sourceEndInput) sourceEndInput.value = toDateValue(sourceEnd);
      if (targetStartInput) targetStartInput.value = toDateValue(monday);
      showAppToast("已填入上周到本周的复制周期。", "success");
    });
  }

  modalOpenButtons.forEach((button) => {
    button.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      openModal(findModalTarget(button.dataset.modalOpen || ""));
    });
  });

  modalCloseButtons.forEach((button) => {
    button.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      closeModal(button.closest("[data-modal]"));
    });
  });

  modalOverlays.forEach((overlay) => {
    overlay.addEventListener("click", (event) => {
      if (event.target !== overlay || overlay.dataset.closeOnOverlay === "0") {
        return;
      }
      closeModal(overlay);
    });
  });

  document.addEventListener("keydown", (event) => {
    if (event.key !== "Escape") {
      return;
    }
    const openModals = Array.from(document.querySelectorAll("[data-modal].is-open"));
    const latestModal = openModals[openModals.length - 1];
    if (latestModal) {
      closeModal(latestModal);
      return;
    }
    const openDrawers = Array.from(document.querySelectorAll("[data-drawer].is-open"));
    const latestDrawer = openDrawers[openDrawers.length - 1];
    if (latestDrawer) {
      closeDrawer(latestDrawer);
    }
  });

  drawerOpenButtons.forEach((button) => {
    button.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      openDrawer(findDrawerTarget(button.dataset.drawerOpen || ""));
    });
  });

  drawerCloseButtons.forEach((button) => {
    button.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      closeDrawer(button.closest("[data-drawer]"));
    });
  });

  drawerOverlays.forEach((drawer) => {
    drawer.addEventListener("click", (event) => {
      if (event.target === drawer) {
        closeDrawer(drawer);
      }
    });
  });

  employeeStoreForms.forEach((form) => {
    const primarySelect = form.querySelector("[data-primary-store-select]");
    if (primarySelect) {
      primarySelect.addEventListener("change", () => syncEmployeePrimaryStoreChoices(form));
    }
    syncEmployeePrimaryStoreChoices(form);
  });

  shiftScopeForms.forEach((form) => {
    Array.from(form.querySelectorAll('input[name="shift_scope"]')).forEach((input) => {
      input.addEventListener("change", () => updateShiftScopeForm(form));
    });
    updateShiftScopeForm(form);
  });

  filterToggleButtons.forEach((button) => {
    button.addEventListener("click", (event) => {
      event.preventDefault();
      const target = document.getElementById(button.dataset.filterToggle || "");
      if (target && target.tagName.toLowerCase() === "details") {
        target.open = !target.open;
      }
    });
  });

  if (rolePermissionToolbar) {
    roleEditorPanels.forEach((panel) => {
      initializeRoleForm(panel.querySelector("[data-role-form]"));
    });
    roleListItems.forEach((item) => {
      item.addEventListener("click", () => {
        switchRolePanel(item.dataset.roleTarget || "");
      });
    });
    roleListFilterButtons.forEach((button) => {
      button.addEventListener("click", () => {
        activeRoleListFilter = button.dataset.roleFilter || "all";
        roleListFilterButtons.forEach((item) => item.classList.toggle("active", item === button));
        updateRoleList();
      });
    });
    if (roleListSearch) {
      roleListSearch.addEventListener("input", updateRoleList);
    }
    roleModuleFilterButtons.forEach((button) => {
      button.addEventListener("click", () => {
        activeRolePermissionModule = button.dataset.roleModuleFilter || "all";
        roleModuleFilterButtons.forEach((item) => item.classList.toggle("active", item === button));
        updateRolePermissionMatrix();
      });
    });
    if (rolePermissionSearch) {
      rolePermissionSearch.addEventListener("input", updateRolePermissionMatrix);
    }
    rolePermissionCards.forEach((card) => {
      card.querySelectorAll("[data-role-permission-checkbox]").forEach((input) => {
        input.addEventListener("change", () => syncRolePermissionCheckboxes(input));
      });
      card.querySelectorAll("[data-permission-module-group]").forEach((group) => {
        updateRolePermissionCounts(group.closest("form"));
      });
      const form = card.querySelector("form");
      const selectModuleButton = card.querySelector("[data-role-select-module]");
      const clearModuleButton = card.querySelector("[data-role-clear-module]");
      if (selectModuleButton) {
        selectModuleButton.addEventListener("click", () => setCurrentRoleModuleChecked(form, true));
      }
      if (clearModuleButton) {
        clearModuleButton.addEventListener("click", () => setCurrentRoleModuleChecked(form, false));
      }
      if (form) {
        form.addEventListener("submit", (event) => {
          const submitter = event.submitter;
          if (submitter && submitter.matches("[data-restore-defaults]")) {
            return;
          }
        });
      }
    });
    window.addEventListener("beforeunload", (event) => {
      if (!hasDirtyRoleForm()) {
        return;
      }
      event.preventDefault();
      event.returnValue = "";
    });
    updateRoleList();
    updateRolePermissionMatrix();
  }

  if (advancedFilterPanel && advancedFilterToggle) {
    advancedFilterToggle.addEventListener("click", (event) => {
      event.preventDefault();
      advancedFilterPanel.open = !advancedFilterPanel.open;
      updateAdvancedFilterToggle();
    });
    advancedFilterPanel.addEventListener("toggle", updateAdvancedFilterToggle);
    updateAdvancedFilterToggle();
  }

  dateRangeSelects.forEach((select) => {
    select.addEventListener("change", updateCustomDateRangeVisibility);
  });
  updateCustomDateRangeVisibility();

  filterDropdowns.forEach((dropdown) => {
    updateFilterDropdownSummary(dropdown);
    snapshotFilterDropdown(dropdown);
    dropdown.addEventListener("toggle", () => {
      if (dropdown.open) {
        snapshotFilterDropdown(dropdown);
        closeOtherFilterDropdowns(dropdown);
        const searchInput = dropdown.querySelector("[data-filter-dropdown-search]");
        if (searchInput) {
          window.setTimeout(() => searchInput.focus(), 0);
        }
      }
    });
    filterDropdownCheckboxes(dropdown).forEach((input) => {
      input.addEventListener("change", () => updateFilterDropdownSummary(dropdown));
    });
    const searchInput = dropdown.querySelector("[data-filter-dropdown-search]");
    if (searchInput) {
      searchInput.addEventListener("input", () => filterDropdownOptionList(dropdown, searchInput.value));
    }
    const selectAllButton = dropdown.querySelector("[data-filter-select-all]");
    if (selectAllButton) {
      selectAllButton.addEventListener("click", (event) => {
        event.preventDefault();
        setVisibleFilterDropdownOptionsChecked(dropdown, true);
      });
    }
    const clearButton = dropdown.querySelector("[data-filter-clear]");
    if (clearButton) {
      clearButton.addEventListener("click", (event) => {
        event.preventDefault();
        setVisibleFilterDropdownOptionsChecked(dropdown, false);
      });
    }
    const confirmButton = dropdown.querySelector("[data-filter-confirm]");
    if (confirmButton) {
      confirmButton.addEventListener("click", (event) => {
        event.preventDefault();
        snapshotFilterDropdown(dropdown);
        updateFilterDropdownSummary(dropdown);
        closeFilterDropdown(dropdown, false);
      });
    }
    const cancelButton = dropdown.querySelector("[data-filter-cancel]");
    if (cancelButton) {
      cancelButton.addEventListener("click", (event) => {
        event.preventDefault();
        closeFilterDropdown(dropdown, true);
      });
    }
  });

  if (filterDropdowns.length) {
    document.addEventListener("click", (event) => {
      if (event.target.closest("[data-filter-dropdown]")) {
        return;
      }
      filterDropdowns.forEach((dropdown) => closeFilterDropdown(dropdown, false));
    });
    document.addEventListener("keydown", (event) => {
      if (event.key === "Escape") {
        filterDropdowns.forEach((dropdown) => closeFilterDropdown(dropdown, false));
      }
    });
  }

  if (dashboardFilterForm) {
    dashboardFilterForm.addEventListener("submit", () => {
      filterDropdowns.forEach((dropdown) => {
        snapshotFilterDropdown(dropdown);
        updateFilterDropdownSummary(dropdown);
      });
    });
  }

  residueCleanupForms.forEach((form) => {
    form.addEventListener("submit", (event) => {
      const selectedTargets = Array.from(form.querySelectorAll('input[name="target"]:checked'));
      const confirmInput = form.querySelector('input[name="confirm_text"]');
      const confirmText = confirmInput ? confirmInput.value.trim() : "";
      const requiredChecks = ["confirm_backup", "confirm_residue", "confirm_irreversible"];
      const allChecksPassed = requiredChecks.every((name) => {
        const field = form.querySelector(`input[name="${name}"]`);
        return field && field.checked;
      });
      if (!selectedTargets.length) {
        event.preventDefault();
        showAppToast("请选择要清理的历史残留。", "error");
        return;
      }
      if (confirmText !== "DELETE_PERSONNEL_RESIDUE" || !allChecksPassed) {
        event.preventDefault();
        showAppToast("请完成历史残留清理的二次确认。", "error");
        return;
      }
      if (!window.confirm("确认执行历史残留清理吗？可硬删除项不可恢复，其余项只会归档/隐藏。")) {
        event.preventDefault();
      }
    });
  });

  confirmForms.forEach((form) => {
    form.addEventListener("submit", (event) => {
      if (event.defaultPrevented || form.matches("[data-residue-cleanup-form]")) {
        return;
      }
      const submitter = event.submitter || document.activeElement;
      const message =
        (submitter && submitter.dataset ? submitter.dataset.confirmMessage : "") || form.dataset.confirmMessage || "";
      if (message && !window.confirm(message)) {
        event.preventDefault();
        event.stopImmediatePropagation();
      }
    });
  });

  feedbackForms.forEach((form) => {
    form.addEventListener("submit", (event) => {
      if (event.defaultPrevented) {
        return;
      }
      const submitter = event.submitter || form.querySelector('[type="submit"]');
      if (!submitter) {
        return;
      }
      submitter.dataset.originalText = submitter.dataset.originalText || submitter.textContent;
      submitter.textContent =
        form.getAttribute("data-loading-label") || submitter.getAttribute("data-loading-label") || "处理中...";
      submitter.disabled = true;
    });
  });

  preserveScrollForms.forEach((form) => {
    form.addEventListener("submit", saveAdminScrollPosition);
  });

  cleanEmptyQueryForms.forEach((form) => {
    form.addEventListener("submit", () => {
      Array.from(form.querySelectorAll('input[name="shift_type_id"], select[name="shift_type_id"]')).forEach((field) => {
        if (!field.value) {
          field.disabled = true;
        }
      });
    });
  });

  restoreAdminScrollPosition();

  document.querySelectorAll(".alert-success, .alert-error").forEach((alert) => {
    const message = alert.textContent ? alert.textContent.trim() : "";
    if (message) {
      showAppToast(message, alert.classList.contains("alert-error") ? "error" : "success");
    }
  });

  if (restShiftButton && scheduleBulkForm) {
    restShiftButton.addEventListener("click", (event) => {
      event.preventDefault();
      const shiftSelect = scheduleBulkForm.querySelector("[data-shift-select]");
      if (!shiftSelect) {
        return;
      }
      const restOption = Array.from(shiftSelect.options).find((option) => (option.dataset.shiftName || option.textContent).includes("休息"));
      if (restOption) {
        shiftSelect.value = restOption.value;
        showAppToast("已选择休息班次，请选择日期和员工后保存。", "success");
      }
    });
  }

  if (copyYesterdayButton) {
    copyYesterdayButton.addEventListener("click", (event) => {
      event.preventDefault();
      showAppToast("请选择目标日期后，按昨日排班模板复制。", "success");
    });
  }

  adminTicketCreateForms.forEach((form) => {
    form.addEventListener("submit", async (event) => {
      event.preventDefault();
      setModalError(form, "");
      const submitButton = event.submitter || form.querySelector('[type="submit"]');
      if (submitButton) {
        submitButton.disabled = true;
      }
      try {
        const response = await fetch(form.action, {
          method: "POST",
          body: new FormData(form),
          credentials: "same-origin",
          headers: {
            Accept: "application/json",
          },
        });
        let payload = {};
        try {
          payload = await response.json();
        } catch (_error) {
          payload = {};
        }
        if (!response.ok || !payload.ok) {
          setModalError(form, payload.error || "工单创建失败，请检查后重试。");
          return;
        }
        closeModal(form.closest("[data-modal]"));
        showAppToast(payload.message || "工单已创建。", "success");
        window.setTimeout(() => {
          window.location.reload();
        }, 650);
      } catch (_error) {
        setModalError(form, "工单创建失败，请检查网络后重试。");
      } finally {
        if (submitButton) {
          submitButton.disabled = false;
        }
      }
    });
  });

  if (notificationRoot) {
    if (notificationDesktopButton && "Notification" in window) {
      notificationDesktopButton.hidden = false;
      if (Notification.permission === "granted" && notificationDesktopEnabled) {
        notificationDesktopButton.textContent = "桌面提醒已开启";
        notificationDesktopButton.disabled = true;
      }
      notificationDesktopButton.addEventListener("click", async () => {
        const permission = await Notification.requestPermission();
        if (permission === "granted") {
          notificationDesktopEnabled = true;
          safeLocalStorageSet("storeRequestDesktopNotifications", "1");
          notificationDesktopButton.textContent = "桌面提醒已开启";
          notificationDesktopButton.disabled = true;
        }
      });
    }

    if (notificationToggle && notificationPanel) {
      notificationToggle.addEventListener("click", (event) => {
        event.preventDefault();
        event.stopPropagation();
        notificationPanel.hidden = !notificationPanel.hidden;
        if (!notificationPanel.hidden) {
          refreshNotifications();
        }
      });
      document.addEventListener("click", (event) => {
        if (!notificationRoot.contains(event.target)) {
          notificationPanel.hidden = true;
        }
      });
    }

    if (notificationList) {
      notificationList.addEventListener("click", async (event) => {
        const actionButton = event.target.closest("[data-notification-action-url]");
        if (actionButton) {
          event.preventDefault();
          const actionUrl = actionButton.dataset.notificationActionUrl || "";
          if (!actionUrl) {
            return;
          }
          await notificationPost(actionUrl);
          await refreshNotifications();
          return;
        }
        const readButton = event.target.closest("[data-notification-read]");
        if (!readButton) {
          return;
        }
        event.preventDefault();
        await notificationPost(`/admin/api/notifications/${readButton.dataset.notificationRead}/read`);
        await refreshNotifications();
      });
    }

    if (notificationReadAll) {
      notificationReadAll.addEventListener("click", async (event) => {
        event.preventDefault();
        await notificationPost("/admin/api/notifications/read-all");
        await refreshNotifications();
      });
    }

    refreshNotifications();
    setInterval(pollNotifications, 15000);
  }

  copyTicketButtons.forEach((button) => {
    button.addEventListener("click", async () => {
      const ticketNo = button.dataset.copyTicket || "";
      if (!ticketNo) {
        return;
      }
      try {
        const copied = await copyText(ticketNo);
        setTemporaryButtonText(button, copied ? "已复制" : "请手动复制");
      } catch (_error) {
        setTemporaryButtonText(button, "请手动复制");
      }
    });
  });
}

document.addEventListener("DOMContentLoaded", () => {
  try {
    initializeStoreRequestApp();
  } catch (error) {
    console.error("Store request UI initialization failed", error);
  }
});
