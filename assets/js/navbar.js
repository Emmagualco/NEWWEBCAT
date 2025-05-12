document.addEventListener("DOMContentLoaded", function () {
  const toggleBtn = document.querySelector(".navbar-toggle");
  const navMenu = document.querySelector(".navbar-nav");

  if (toggleBtn && navMenu) {
    toggleBtn.addEventListener("click", () => {
      navMenu.classList.toggle("active");
      toggleBtn.classList.toggle("open");
    });
  }
});

  