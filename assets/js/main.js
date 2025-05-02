// =============== Hero Slider ===================
let currentSlide = 0;
const slides = document.querySelectorAll('#hero .slide');

function showNextSlide() {
  slides[currentSlide].classList.remove('active');
  currentSlide = (currentSlide + 1) % slides.length;
  slides[currentSlide].classList.add('active');
}

if (slides.length > 0) {
  setInterval(showNextSlide, 5000);
}

// =============== Contact Form Validation ===================
const contactForm = document.querySelector('.contact-form');
if (contactForm) {
  contactForm.addEventListener('submit', function (e) {
    const nombre = document.getElementById('nombre').value.trim();
    const email = document.getElementById('email').value.trim();
    const mensaje = document.getElementById('mensaje').value.trim();
    const emailRegex = /^[^\\s@]+@[^\\s@]+\\.[^\\s@]+$/;

    if (!nombre || !email || !mensaje) {
      alert('Por favor, completa todos los campos.');
      e.preventDefault();
    } else if (!emailRegex.test(email)) {
      alert('El correo electrónico no es válido.');
      e.preventDefault();
    }
  });
}

// =============== Igualar Alturas de Cards (Servicios) ===================
function equalHeight() {
  const cards = document.querySelectorAll('.card');
  let max = 0;
  cards.forEach(card => {
    card.style.height = 'auto';
    if (card.offsetHeight > max) max = card.offsetHeight;
  });
  cards.forEach(card => card.style.height = max + 'px');
}

window.addEventListener('load', equalHeight);
window.addEventListener('resize', equalHeight);

const swiper = new Swiper('.mySwiper', {
  slidesPerView: 4,
  spaceBetween: 30,
  loop: true,
  autoplay: {
    delay: 2500,
    disableOnInteraction: false,
  },
  breakpoints: {
    320: {
      slidesPerView: 2,
      spaceBetween: 20
    },
    768: {
      slidesPerView: 3,
      spaceBetween: 30
    },
    1024: {
      slidesPerView: 5,
      spaceBetween: 50
    }
  }
});



