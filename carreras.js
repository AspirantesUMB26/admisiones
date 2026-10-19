// carreras.js (corregido y con exposición global)
const DB_CARRERAS = {
  "ingenieria-mecanica": {
    titulo: "Ingeniería Mecánica",
    duracion: "9 semestres",
    planteles: ["UES Santiago Tianguistenco"],
    heroImg: "assets/images/hero-mecanica.jpg",
    objetivo: "Formar profesionales capaces de diseñar, analizar, operar y mantener sistemas mecánicos y térmicos.",
    perfilIngreso: [
      "Afinidad por la física y matemáticas",
      "Interés por el funcionamiento de máquinas",
      "Dibujo técnico"
    ],
    perfilEgreso: [
      "Diseña sistemas de transmisión de potencia",
      "Mantiene equipo industrial de precisión"
    ],
    planEstudios: [
      {
        semestre: "1er Semestre",
        materias: ["Cálculo Diferencial", "Química", "Taller de Ética", "Dibujo Mecánico", "Fundamentos de Investigación", "Inglés I"]
      },
      {
        semestre: "2do Semestre",
        materias: ["Cálculo Integral", "Estática", "Algoritmos y Programación", "Metrología", "Comunicación Profesional", "Inglés II"]
      },
      {
        semestre: "3er Semestre",
        materias: ["Cálculo Vectorial", "Resistencia de Materiales", "Termodinámica", "Diseño Asistido por Computadora", "Inglés III"]
      }
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
    perfilIngreso: [
      "Interés por el cuidado de la salud",
      "Habilidades de comunicación",
      "Responsabilidad y ética profesional"
    ],
    perfilEgreso: [
      "Brinda atención integral en salud",
      "Aplica protocolos de enfermería",
      "Participa en equipos interdisciplinarios"
    ],
    planEstudios: [
      {
        semestre: "1er Semestre",
        materias: ["Anatomía Humana", "Bioquímica", "Fundamentos de Enfermería", "Comunicación", "Inglés I"]
      },
      {
        semestre: "2do Semestre",
        materias: ["Fisiología Humana", "Microbiología", "Nutrición", "Enfermería Básica", "Inglés II"]
      },
      {
        semestre: "3er Semestre",
        materias: ["Patología General", "Farmacología", "Enfermería del Adulto", "Salud Pública", "Inglés III"]
      }
    ],
    campoLaboral: [
      { area: "Hospitalaria", desc: "Atención y cuidado de pacientes" },
      { area: "Salud Pública", desc: "Programas de prevención y promoción" }
    ],
    contacto: { tel: "(01 713) 135 2297", email: "informes@umb.mx" }
  },
  "licenciatura-en-administracion": {
    titulo: "Licenciatura en Administración",
    duracion: "9 semestres",
    planteles: ["UES Coatepec Harinas", "UES Jilotepec"],
    heroImg: "assets/images/hero-administracion.jpg",
    objetivo: "Desarrollar líderes con capacidad para administrar recursos, procesos y organizaciones con enfoque estratégico y responsable.",
    perfilIngreso: [
      "Interés por la gestión y la toma de decisiones",
      "Capacidad para trabajar en equipo",
      "Comprensión de procesos operativos"
    ],
    perfilEgreso: [
      "Administra recursos y proyectos",
      "Participa en la planeación estratégica de organizaciones"
    ],
    planEstudios: [
      {
        semestre: "1er Semestre",
        materias: ["Introducción a la Administración", "Matemáticas", "Contabilidad", "Comunicación", "Inglés I"]
      },
      {
        semestre: "2do Semestre",
        materias: ["Microeconomía", "Estadística", "Mercadotecnia", "Organización Empresarial", "Inglés II"]
      }
    ],
    campoLaboral: [
      { area: "Empresas", desc: "Administración de áreas y procesos" },
      { area: "Consultoría", desc: "Diagnóstico y mejora de operaciones" }
    ],
    contacto: { tel: "(01 713) 135 2297", email: "informes@umb.mx" }
  }
};

if (typeof window !== 'undefined') {
  window.DB_CARRERAS = DB_CARRERAS;
}
