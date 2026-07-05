document.addEventListener("DOMContentLoaded", () => {
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
  const scheduleBulkForm = document.querySelector("[data-schedule-bulk-form]");
  const scheduleBulkSummary = document.querySelector("[data-schedule-bulk-summary]");
  const selectAllEmployeesButton = document.querySelector("[data-select-all-employees]");
  const clearEmployeesButton = document.querySelector("[data-clear-employees]");
  const selectAllScheduleDatesButton = document.querySelector("[data-select-all-schedule-dates]");
  const selectWeekdaysButton = document.querySelector("[data-select-weekdays]");
  const selectWeekendsButton = document.querySelector("[data-select-weekends]");
  const clearScheduleDatesButton = document.querySelector("[data-clear-schedule-dates]");

  let selectedImages = [];
  let selectedFiles = [];
  let selectedExtraImages = [];
  let selectedExtraFiles = [];
  let notificationLatestId = 0;
  let notificationInitialLoaded = false;
  let notificationDesktopEnabled = window.localStorage.getItem("storeRequestDesktopNotifications") === "1";

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
  }

  function closeModal(modal) {
    if (!modal) {
      return;
    }
    modal.classList.remove("is-open");
    modal.hidden = true;
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

  function setScheduleInputsChecked(inputs, checked) {
    inputs.forEach((input) => {
      input.checked = checked;
    });
    updateScheduleBulkSummary();
  }

  function setScheduleDatesByWeekend(weekendOnly) {
    scheduleDateInputs().forEach((input) => {
      const chip = input.closest("[data-weekend]");
      const isWeekend = chip ? chip.dataset.weekend === "1" : false;
      input.checked = weekendOnly ? isWeekend : !isWeekend;
    });
    updateScheduleBulkSummary();
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
    requestTypeHint.textContent = selectedRule.description_hint || "不同需求类型可能要求补充品牌、商品、数量或附件。";
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
    scheduleEmployeeInputs().forEach((input) => input.addEventListener("change", updateScheduleBulkSummary));
    scheduleDateInputs().forEach((input) => input.addEventListener("change", updateScheduleBulkSummary));
    if (selectAllEmployeesButton) {
      selectAllEmployeesButton.addEventListener("click", () => setScheduleInputsChecked(scheduleEmployeeInputs(), true));
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
          window.localStorage.setItem("storeRequestDesktopNotifications", "1");
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
});
