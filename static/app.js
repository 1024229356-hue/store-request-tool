document.addEventListener("DOMContentLoaded", () => {
  const imageInput = document.querySelector("[data-image-input]");
  const imagePreview = document.querySelector("[data-image-preview]");
  const clearImagesButton = document.querySelector("[data-clear-images]");
  const fileInput = document.querySelector("[data-file-input]");
  const filePreview = document.querySelector("[data-file-preview]");
  const clearFilesButton = document.querySelector("[data-clear-files]");
  const copyTicketButtons = document.querySelectorAll("[data-copy-ticket]");

  let selectedImages = [];
  let selectedFiles = [];

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
