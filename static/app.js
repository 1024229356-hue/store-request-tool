(function () {
  function formatSize(bytes) {
    if (bytes >= 1024 * 1024) {
      return `${(bytes / (1024 * 1024)).toFixed(1)}MB`;
    }
    if (bytes >= 1024) {
      return `${(bytes / 1024).toFixed(1)}KB`;
    }
    return `${bytes}B`;
  }

  function replaceInputFiles(input, files) {
    if (typeof DataTransfer === "undefined") {
      return;
    }
    const transfer = new DataTransfer();
    files.forEach((file) => transfer.items.add(file));
    input.files = transfer.files;
  }

  function setupUploadList(options) {
    const input = document.querySelector(options.inputSelector);
    const preview = document.querySelector(options.previewSelector);
    const clearButton = document.querySelector(options.clearSelector);
    if (!input || !preview) {
      return;
    }

    let selectedFiles = Array.from(input.files || []);

    function syncFiles(nextFiles) {
      selectedFiles = nextFiles;
      replaceInputFiles(input, selectedFiles);
      render();
    }

    function render() {
      preview.innerHTML = "";
      selectedFiles.forEach((file, index) => {
        const item = document.createElement("div");
        item.className = options.image ? "preview-item preview-item-image" : "preview-item preview-item-file";

        if (options.image && file.type.startsWith("image/")) {
          const image = document.createElement("img");
          image.alt = file.name;
          image.src = URL.createObjectURL(file);
          image.onload = () => URL.revokeObjectURL(image.src);
          item.appendChild(image);
        }

        const meta = document.createElement("div");
        meta.className = "preview-meta";

        const name = document.createElement("span");
        name.className = "preview-name";
        name.textContent = file.name;
        meta.appendChild(name);

        const size = document.createElement("span");
        size.className = "preview-size";
        size.textContent = formatSize(file.size);
        meta.appendChild(size);

        item.appendChild(meta);

        const remove = document.createElement("button");
        remove.className = "ghost-button compact preview-remove";
        remove.type = "button";
        remove.textContent = "删除";
        remove.addEventListener("click", () => {
          syncFiles(selectedFiles.filter((_file, fileIndex) => fileIndex !== index));
        });
        item.appendChild(remove);

        preview.appendChild(item);
      });
    }

    input.addEventListener("change", () => {
      syncFiles(Array.from(input.files || []));
    });

    if (clearButton) {
      clearButton.addEventListener("click", () => {
        syncFiles([]);
      });
    }

    render();
  }

  try {
    setupUploadList({
      inputSelector: "[data-image-input]",
      previewSelector: "[data-image-preview]",
      clearSelector: "[data-image-clear]",
      image: true,
    });
    setupUploadList({
      inputSelector: "[data-file-input]",
      previewSelector: "[data-file-preview]",
      clearSelector: "[data-file-clear]",
      image: false,
    });
  } catch (_error) {
    // Form submission must keep working even if preview management is unavailable.
  }
})();
