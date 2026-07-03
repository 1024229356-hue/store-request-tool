const input = document.querySelector("[data-image-input]");
const preview = document.querySelector("[data-image-preview]");

if (input && preview) {
  input.addEventListener("change", () => {
    preview.innerHTML = "";
    Array.from(input.files || []).forEach((file) => {
      const item = document.createElement("div");
      item.className = "preview-item";
      const name = document.createElement("span");
      name.textContent = file.name;
      item.appendChild(name);

      if (file.type.startsWith("image/")) {
        const image = document.createElement("img");
        image.alt = file.name;
        image.src = URL.createObjectURL(file);
        image.onload = () => URL.revokeObjectURL(image.src);
        item.prepend(image);
      }
      preview.appendChild(item);
    });
  });
}
