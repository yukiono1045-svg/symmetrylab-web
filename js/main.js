/* ============================================
   SYMMETRY Lab - Main JavaScript
   ============================================ */

// Mobile Navigation Toggle
document.addEventListener('DOMContentLoaded', () => {
  const toggle = document.querySelector('.nav-toggle');
  const links = document.querySelector('.nav-links');

  if (toggle && links) {
    toggle.addEventListener('click', () => {
      links.classList.toggle('active');
      toggle.setAttribute('aria-expanded', links.classList.contains('active'));
    });

    // Close menu when clicking a link
    links.querySelectorAll('a').forEach(link => {
      link.addEventListener('click', () => {
        links.classList.remove('active');
        toggle.setAttribute('aria-expanded', 'false');
      });
    });
  }

  // Scroll-triggered fade-in animations
  const observer = new IntersectionObserver((entries) => {
    entries.forEach(entry => {
      if (entry.isIntersecting) {
        entry.target.classList.add('visible');
        observer.unobserve(entry.target);
      }
    });
  }, { threshold: 0.1 });

  document.querySelectorAll('.fade-in').forEach(el => observer.observe(el));

  // Contact form validation
  const contactForm = document.querySelector('#contact-form');
  if (contactForm) {
    contactForm.addEventListener('submit', (e) => {
      const name = contactForm.querySelector('[name="name"]');
      const email = contactForm.querySelector('[name="email"]');
      let valid = true;

      if (name && !name.value.trim()) {
        valid = false;
        name.style.borderColor = '#EF4444';
      } else if (name) {
        name.style.borderColor = '';
      }

      if (email && !email.value.match(/^[^\s@]+@[^\s@]+\.[^\s@]+$/)) {
        valid = false;
        email.style.borderColor = '#EF4444';
      } else if (email) {
        email.style.borderColor = '';
      }

      if (!valid) {
        e.preventDefault();
        alert('必須項目を正しく入力してください。');
      }
    });
  }

  // Booking success/cancel handling
  const urlParams = new URLSearchParams(window.location.search);
  const bookingMsg = document.querySelector('#booking-message');
  if (bookingMsg) {
    if (urlParams.get('success') === 'true') {
      bookingMsg.innerHTML = '<div class="alert alert-success">お申し込みありがとうございます。確認メールをお送りしましたのでご確認ください。</div>';
    } else if (urlParams.get('canceled') === 'true') {
      bookingMsg.innerHTML = '<div class="alert alert-warning">お支払いがキャンセルされました。再度お試しください。</div>';
    }
  }
});
