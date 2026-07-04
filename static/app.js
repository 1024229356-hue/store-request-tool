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

  let selectedImages = [];
  let selectedFiles = [];
  let selectedExtraImages = [];
  let selectedExtraFiles = [];

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
