document.addEventListener("DOMContentLoaded", function () {
  const menuToggle = document.querySelector('.navbar-toggle');
  const mobileMenu = document.querySelector('.navbar-nav');
  
  if (menuToggle) {
    menuToggle.addEventListener('click', function() {
      mobileMenu.classList.toggle('active');
      const isExpanded = mobileMenu.classList.contains('active');
      menuToggle.setAttribute('aria-expanded', isExpanded);
    });
  }
});



  