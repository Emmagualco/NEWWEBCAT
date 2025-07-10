document.addEventListener("DOMContentLoaded", () => {
  const container = document.getElementById("timelineContainer");
  const btnPrev = document.getElementById("btnPrev");
  const btnNext = document.getElementById("btnNext");
  const dotsContainer = document.getElementById("timelineDots");

  const items = container.querySelectorAll(".timeline-item");

  function getItemWidth() {
    const style = window.getComputedStyle(items[0]);
    const marginRight = parseInt(style.marginRight) || 0;
    return items[0].offsetWidth + marginRight;
  }

  let currentIndex = 0;

  // Crear los dots dinámicamente
  items.forEach((_, index) => {
    const dot = document.createElement("button");
    dot.classList.add("timeline-dot");
    if (index === 0) dot.classList.add("active");
    dot.addEventListener("click", () => {
      const itemWidth = getItemWidth();
      container.scrollTo({ left: index * itemWidth, behavior: "smooth" });
      updateDots(index);
      currentIndex = index;
    });
    dotsContainer.appendChild(dot);
  });

  function updateDots(index) {
    const dots = dotsContainer.querySelectorAll(".timeline-dot");
    dots.forEach(dot => dot.classList.remove("active"));
    if (dots[index]) dots[index].classList.add("active");
  }

  btnPrev.addEventListener("click", () => {
    if (currentIndex > 0) currentIndex--;
    const itemWidth = getItemWidth();
    container.scrollTo({ left: currentIndex * itemWidth, behavior: "smooth" });
    updateDots(currentIndex);
  });

  btnNext.addEventListener("click", () => {
    if (currentIndex < items.length - 1) currentIndex++;
    const itemWidth = getItemWidth();
    container.scrollTo({ left: currentIndex * itemWidth, behavior: "smooth" });
    updateDots(currentIndex);
  });

  // ✅ Centrar la card destacada al ingresar
  const destacado = document.getElementById("inicio-slider");
  if (destacado) {
    const destacadoIndex = Array.from(items).indexOf(destacado);
    if (destacadoIndex !== -1) {
      currentIndex = destacadoIndex;
      updateDots(currentIndex);

      const containerWidth = container.offsetWidth;
      const itemCenterOffset = destacado.offsetLeft + (destacado.offsetWidth / 2);
      const scrollPosition = itemCenterOffset - (containerWidth / 2);

      // Scroll para centrar la card destacado
      container.scrollTo({
        left: scrollPosition,
        behavior: "smooth"
      });

      // Iniciar el auto-slide después de 1 segundo (para que termine el scroll smooth)
      setTimeout(() => {
        startAutoSlide();
      }, 1000);
    } else {
      // Si no se encontró destacado, iniciar auto-slide normal
      startAutoSlide();
    }
  } else {
    // Si no hay destacado, iniciar auto-slide normal
    startAutoSlide();
  }

  // Auto-slide (activo)
  function startAutoSlide() {
    setInterval(() => {
      if (currentIndex < items.length - 1) {
        currentIndex++;
      } else {
        currentIndex = 0;
      }
      const itemWidth = getItemWidth();
      container.scrollTo({ left: currentIndex * itemWidth, behavior: "smooth" });
      updateDots(currentIndex);
    }, 5000);
  }
});





