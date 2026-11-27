if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('./service-worker.js').catch((error) => {
      console.warn('PWA belum dapat diaktifkan:', error);
    });
  });
}

// WhatsApp Web membaca Unicode lebih konsisten melalui endpoint send.
document.querySelectorAll('a[href*="wa.me/"]').forEach((link) => {
  const source = new URL(link.href);
  const phone = source.pathname.replace(/\D/g, '');
  const message = source.searchParams.get('text') || '';
  const target = new URL('https://api.whatsapp.com/send');
  target.searchParams.set('phone', phone);
  target.searchParams.set('text', message);
  link.href = target.href;
});
