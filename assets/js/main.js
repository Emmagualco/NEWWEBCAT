// =============== Hero Slider ===================
document.addEventListener("DOMContentLoaded", () => {
  let currentSlide = 0;
  const slides = document.querySelectorAll('#hero .slide');

  if (slides.length > 0) {
    slides[currentSlide].classList.add('active');

    setInterval(() => {
      slides[currentSlide].classList.remove('active');
      currentSlide = (currentSlide + 1) % slides.length;
      slides[currentSlide].classList.add('active');
    }, 5000);
  }

  // =============== Igualar Alturas de Cards (solo en escritorio) ===================
  function equalHeight() {
    const cards = document.querySelectorAll('.card');

    // Reset heights
    cards.forEach(card => {
      card.style.height = 'auto';
    });

    // Solo aplicar en pantallas grandes
    if (window.innerWidth > 768) {
      let max = 0;
      cards.forEach(card => {
        if (card.offsetHeight > max) max = card.offsetHeight;
      });
      cards.forEach(card => card.style.height = max + 'px');
    }
  }

  window.addEventListener('load', equalHeight);
  window.addEventListener('resize', equalHeight);

  // =============== Swiper Carrusel ===================
  if (document.querySelector('.mySwiper')) {
    const swiper = new Swiper('.mySwiper', {
      loop: true,
      autoplay: {
        delay: 2500,
        disableOnInteraction: false,
      },
      slidesPerView: 5,
      spaceBetween: 30,
      breakpoints: {
        320: {
          slidesPerView: 2,
          spaceBetween: 10
        },
        480: {
          slidesPerView: 3,
          spaceBetween: 15
        },
        768: {
          slidesPerView: 4,
          spaceBetween: 20
        },
        1024: {
          slidesPerView: 5,
          spaceBetween: 30
        }
      }
    });
  }

  // =============== Navbar Toggle Responsive ===================
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





