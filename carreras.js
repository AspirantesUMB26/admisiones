// carreras.js (corregido y con exposición global)
const DB_CARRERAS = {
  "ingenieria-mecanica": {
    titulo: "Ingeniería Mecánica",
    duracion: "9 semestres",
    planteles: ["UES Santiago Tianguistenco"],
    heroImg: "assets/images/hero-mecanica.jpg",
    objetivo: "Formar profesionales capaces de diseñar, analizar, operar y mantener sistemas mecánicos y térmicos.",
    perfilIngreso: ["Afinidad por la física y matemáticas", "Interés por el funcionamiento de máquinas", "Dibujo técnico"],
    perfilEgreso: ["Diseña sistemas de transmisión de potencia", "Mantiene equipo industrial de precisión"],
    planEstudios: [
      { semestre: "1er Semestre", materias: ["Cálculo Diferencial", "Química", "Taller de Ética", "Dibujo Mecánico", "Fundamentos de Investigación", "Inglés I"] },
      { semestre: "2do Semestre", materias: ["Cálculo Integral", "Estática", "Algoritmos y Programación", "Metrología", "Comunicación Profesional", "Inglés II"] }
      // ... resto de semestres (agregar según plan real)
    ],
    campoLaboral: [
      { area: "Metalmecánica", desc: "Diseño de moldes y troqueles" },
      { area: "Plantas Industriales", desc: "Mantenimiento y optimización de procesos" }
    ],
    contacto: { tel: "(01 713) 135 2297", email: "informes@umb.mx" }
  },
  "licenciatura-en-enfermeria": {
    titulo: "Licenciatura en Enfermería",
    duracion: "10 semestres",
    planteles: ["UES Coatepec Harinas", "UES El Oro", "UES Jilotepec", "UES Sultepec", "UES Temoaya"],
    heroImg: "assets/images/hero-enfermeria.jpg",
    objetivo: "Formar Licenciados en Enfermería con preparación científica y humanista para la atención primaria y hospitalaria.",
    perfilIngreso: ["Interés por el cuidado de la salud", "Habilidades de comunicación", "Responsabilidad y ética profesional"],
    perfilEgreso: ["Brinda atención integral en salud", "Aplica protocolos de enfermería", "Participa en equipos interdisciplinarios"],
    planEstudios: [
      { semestre: "1er Semestre", materias: ["Anatomía Humana", "Bioquímica", "Fundamentos de Enfermería", "Comunicación", "Inglés I"] }
      // ... completar el plan real
    ],
    campoLaboral: [
      { area: "Hospitalaria", desc: "Atención y cuidado de pacientes" },
      { area: "Salud Pública", desc: "Programas de prevención y promoción" }
    ],
    contacto: { tel: "(01 713) 135 2297", email: "informes@umb.mx" }
  }
  // Agregar más carreras aquí siguiendo la misma estructura.
};

// Exponer globalmente para que carrera.html e index.html lo consuman sin errores
if (typeof window !== 'undefined') {
  window.DB_CARRERAS = DB_CARRERAS;
}
