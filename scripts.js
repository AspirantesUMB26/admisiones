// scripts.js (corregido)
async function downloadInfografia() {
    // 1. Carga automática de la librería si falta
    if (typeof html2canvas === 'undefined') {
        const script = document.createElement('script');
        script.src = "https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js";
        document.head.appendChild(script);
        await new Promise((resolve, reject) => {
            script.onload = resolve;
            script.onerror = () => reject(new Error('No se pudo cargar html2canvas desde CDN'));
        });
    }

    // 2. Obtener el nombre de la carrera
    const h1Element = document.querySelector('h1');
    const careerName = h1Element ? h1Element.innerText.trim() : (document.title.split('|')[0] || 'carrera').trim();
    const safeFileName = careerName.replace(/[^a-z0-9-_]/gi, '_');

    // 3. Elemento a capturar (preferencia: .main-column, .container, body)
    const element = document.querySelector('.main-column') || document.querySelector('.container') || document.body;

    // 4. Pantalla de carga (overlay)
    const loading = document.createElement('div');
    loading.id = 'download-loading-overlay';
    loading.style.cssText = [
        'position:fixed',
        'top:0',
        'left:0',
        'width:100%',
        'height:100%',
        'background:rgba(0,0,0,0.85)',
        'color:white',
        'display:flex',
        'flex-direction:column',
        'justify-content:center',
        'align-items:center',
        'z-index:99999',
        'padding:20px'
    ].join(';');
    loading.innerHTML = '<div style="border:4px solid rgba(255,255,255,0.15);border-top:4px solid #85182a;border-radius:50%;width:48px;height:48px;animation:spin 1s linear infinite;"></div><p style="margin-top:16px;font-size:1rem;">Generando imagen, por favor espera...</p>';
    // añadir pequeña animación CSS si hace falta
    const style = document.createElement('style');
    style.id = 'download-loading-style';
    style.textContent = '@keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }';
    document.head.appendChild(style);
    document.body.appendChild(loading);

    // 5. Crear encabezado temporal dentro de la imagen
    const badge = document.createElement('div');
    badge.id = "temp-badge";
    badge.innerText = careerName;
    badge.style.cssText = [
        'width:100%',
        'background:#85182a',
        'color:white',
        'text-align:center',
        'padding:18px 0',
        'font-size:20px',
        'font-weight:700',
        'text-transform:uppercase',
        'letter-spacing:0.5px',
        'box-sizing:border-box'
    ].join(';');

    // Prepend solo si element existe
    try {
        if (element && element.prepend) {
            element.prepend(badge);
        } else {
            document.body.insertBefore(badge, document.body.firstChild);
        }

        // Forzar visibilidad de elementos animados
        const animatedElements = document.querySelectorAll('.card, .semestre, .fade-in-up');
        animatedElements.forEach(el => {
            el.style.opacity = '1';
            el.style.visibility = 'visible';
            el.style.transform = 'none';
        });

        // 6. Generar canvas con html2canvas
        const canvas = await html2canvas(element, {
            scale: 2,
            useCORS: true,
            backgroundColor: "#ffffff",
            windowWidth: Math.max(document.documentElement.clientWidth || 0, 1200)
        });

        // 7. Descargar la imagen
        const link = document.createElement('a');
        link.href = canvas.toDataURL("image/png");
        link.download = `Infografia_${safeFileName}.png`;
        document.body.appendChild(link);
        link.click();
        link.remove();
    } catch (error) {
        console.error('Error generando infografía:', error);
        alert('No fue posible generar la infografía. Revisa la consola para más detalles.');
    } finally {
        // Limpieza: quitar el título temporal y el overlay
        const tmp = document.getElementById("temp-badge");
        if (tmp && tmp.parentNode) tmp.parentNode.removeChild(tmp);
        const overlay = document.getElementById('download-loading-overlay');
        if (overlay && overlay.parentNode) overlay.parentNode.removeChild(overlay);
        const st = document.getElementById('download-loading-style');
        if (st && st.parentNode) st.parentNode.removeChild(st);
    }
}

window.downloadPDF = downloadInfografia;
