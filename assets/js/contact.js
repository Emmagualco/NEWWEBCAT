// Contact Page JavaScript

// Form validation
document.addEventListener('DOMContentLoaded', function() {
  const form = document.querySelector('.contact-form');
  const inputs = form.querySelectorAll('input, textarea');
  
  // Add real-time validation
  inputs.forEach(input => {
    input.addEventListener('blur', validateField);
    input.addEventListener('input', clearError);
  });
  
  form.addEventListener('submit', handleSubmit);
  
  function validateField(e) {
    const field = e.target;
    const errorElement = document.getElementById(field.name + '-error');
    
    if (field.hasAttribute('required') && !field.value.trim()) {
      showError(field, 'Este campo es requerido');
    } else if (field.type === 'email' && !isValidEmail(field.value)) {
      showError(field, 'Por favor ingrese un email válido');
    } else {
      clearError(field);
    }
  }
  
  function showError(field, message) {
    const errorElement = document.getElementById(field.name + '-error');
    field.classList.add('error');
    if (errorElement) {
      errorElement.textContent = message;
      errorElement.style.display = 'block';
    }
  }
  
  function clearError(field) {
    if (typeof field === 'object' && field.target) {
      field = field.target;
    }
    const errorElement = document.getElementById(field.name + '-error');
    field.classList.remove('error');
    if (errorElement) {
      errorElement.textContent = '';
      errorElement.style.display = 'none';
    }
  }
  
  function isValidEmail(email) {
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(email);
  }
  
  function handleSubmit(e) {
    e.preventDefault();
    
    let isValid = true;
    inputs.forEach(input => {
      if (input.hasAttribute('required') && !input.value.trim()) {
        showError(input, 'Este campo es requerido');
        isValid = false;
      } else if (input.type === 'email' && input.value && !isValidEmail(input.value)) {
        showError(input, 'Por favor ingrese un email válido');
        isValid = false;
      }
    });
    
    if (isValid) {
      const submitBtn = form.querySelector('.btn-submit');
      submitBtn.textContent = 'Enviando...';
      submitBtn.disabled = true;

      // Enviar via fetch a handler server-side
      const formData = new FormData(form);

      fetch('/send_contact.asp', {
        method: 'POST',
        body: formData,
      })
      .then(async (res) => {
        const text = await res.text();
        let data = null;
        try { data = JSON.parse(text); } catch (err) { data = null; }

        if (!res.ok) {
          const msg = (data && data.error) ? data.error : 'Error en el servidor al enviar el mensaje';
          throw new Error(msg);
        }

        if (data && data.success) {
          // mostrar mensaje amigable
          const successMsg = document.createElement('div');
          successMsg.className = 'form-success';
          successMsg.setAttribute('role','status');
          successMsg.textContent = 'Mensaje enviado exitosamente. Te contactaremos pronto.';
          form.parentNode.insertBefore(successMsg, form);
          form.reset();
        } else {
          const msg = (data && data.error) ? data.error : 'No se pudo enviar el mensaje.';
          throw new Error(msg);
        }
      })
      .catch(err => {
        alert('Error: ' + err.message);
      })
      .finally(() => {
        submitBtn.textContent = 'Enviar Mensaje';
        submitBtn.disabled = false;
      });
    }
  }
  
  // Mobile menu toggle
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

// Lazy loading for maps
const mapObserver = new IntersectionObserver((entries) => {
  entries.forEach(entry => {
    if (entry.isIntersecting) {
      const iframe = entry.target;
      if (iframe.dataset.src) {
        iframe.src = iframe.dataset.src;
        iframe.removeAttribute('data-src');
        mapObserver.unobserve(iframe);
      }
    }
  });
});

document.querySelectorAll('iframe[data-src]').forEach(iframe => {
  mapObserver.observe(iframe);
});
  