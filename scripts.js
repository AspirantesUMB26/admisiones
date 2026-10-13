async function downloadInfografia() {
    // 1. Carga automática de la librería si falta
    if (typeof html2canvas === 'undefined') {
        const script = document.createElement('script');
        script.src = "https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js";
        document.head.appendChild(script);
        await new Promise((resolve) => script.onload = resolve);
    }

    // 2. Obtener el nombre de la carrera
    const h1Element = document.querySelector('h1');
    const careerName = h1Element ? h1Element.innerText : document.title.split('|')[0].trim();
    const safeFileName = careerName.replace(/[^a-z0-9]/gi, '_');

    const element = document.querySelector('.main-column') || document.querySelector('.container') || document.body;

    // 3. Pantalla de carga
    const loading = document.createElement('div');
    loading.style.cssText = 'position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.9);color:white;display:flex;justify-content:center;align-items:center;z-index:99999;flex-direction:column;font-family:sans-serif;';
    loading.innerHTML = `<div style="border:4px solid #f3f3f3;border-top:4px solid #85182a;border-radius:50%;width:50px;height:50px;animation:spin 1s linear infinite;"></div><p style="margin-top:20px;">Preparando infografía de ${careerName}...</p><style>@keyframes spin{0%{transform:rotate(0deg)}100%{transform:rotate(360deg)}}</style>`;
    document.body.appendChild(loading);

    // 4. Crear el encabezado visual dentro de la imagen
    const badge = document.createElement('div');
    badge.id = "temp-badge";
    badge.innerText = careerName;
    badge.style.cssText = `
        width: 100%; 
        background: #85182a; 
        color: white; 
        text-align: center; 
        padding: 20px 0; 
        font-size: 24px; 
        font-weight: bold; 
        text-transform: uppercase;
        margin-bottom: 20px;
    `;
    element.prepend(badge); // Lo pone al principio de la imagen

    // 5. Forzar visibilidad
    const animatedElements = document.querySelectorAll('.card, .semestre, .fade-in-up');
    animatedElements.forEach(el => {
        el.style.opacity = '1';
        el.style.visibility = 'visible';
        el.style.transform = 'none';
    });

    try {
        const canvas = await html2canvas(element, {
            scale: 2,
            useCORS: true,
            backgroundColor: "#ffffff",
            windowWidth: 1200
        });

        const link = document.createElement('a');
        link.href = canvas.toDataURL("image/png");
        link.download = `Infografia_${safeFileName}.png`;
        link.click();
    } catch (error) {
        console.error(error);
    } finally {
        // Limpieza: quitar el título temporal para que no se quede en la web
        if (document.getElementById("temp-badge")) badge.remove();
        document.body.removeChild(loading);
    }
}

window.downloadPDF = downloadInfografia;