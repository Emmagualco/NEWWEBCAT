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

  // =============== Contact Form Validation ===================
  const contactForm = document.querySelector('.contact-form');
  if (contactForm) {
    contactForm.addEventListener('submit', function (e) {
      const nombre = document.getElementById('nombre')?.value.trim();
      const email = document.getElementById('email')?.value.trim();
      const mensaje = document.getElementById('mensaje')?.value.trim();
      const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

      if (!nombre || !email || !mensaje) {
        alert('Por favor, completa todos los campos.');
        e.preventDefault();
        return;
      }

      if (!emailRegex.test(email)) {
        alert('El correo electrónico no es válido.');
        e.preventDefault();
      }
    });
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
  const toggleBtn = document.querySelector(".navbar-toggle");
  const navMenu = document.querySelector(".navbar-nav");

  if (toggleBtn && navMenu) {
    toggleBtn.addEventListener("click", () => {
      navMenu.classList.toggle("open");
      toggleBtn.classList.toggle("open");
    });
  }

  // =============== Leaflet Map ===================
  const mapElement = document.getElementById('map');
  if (mapElement) {
    const map = L.map('map', { minZoom: 2 }).setView([-38.4161, -63.6167], 4);

    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
      maxZoom: 10,
      attribution: '© OpenStreetMap contributors'
    }).addTo(map);
  }
});





