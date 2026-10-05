const { obtenerEmbeddings } = require("./bedrock");
const { construirDendrograma, cortarEnGrupos } = require("./clustering");
const { nombrarImagen } = require("./openai");

const UMBRAL_DISTANCIA = Number(process.env.EMBEDDINGS_UMBRAL_DISTANCIA || 0.3);

// Misma interfaz que clasificarImagenes(imagenes) de lib/openai.js: recibe
// [{buffer, extension}, ...] y devuelve un array de categorias (string), una
// por imagen, en el mismo orden de entrada. Asi es un reemplazo directo.
async function clasificarPorEmbeddings(imagenes) {
  console.log(`[ordenar-excel] obteniendo embeddings de ${imagenes.length} imagenes via Bedrock...`);
  const embeddings = await obtenerEmbeddings(imagenes);

  console.log("[ordenar-excel] agrupando por clustering jerarquico...");
  const dendrograma = construirDendrograma(embeddings);
  const grupos = cortarEnGrupos(dendrograma, UMBRAL_DISTANCIA);
  console.log(`[ordenar-excel] ${grupos.length} grupo(s) formados por distancia`);

  const categoriaPorIndice = new Array(imagenes.length);

  for (let g = 0; g < grupos.length; g++) {
    const miembros = grupos[g];
    const imagenRepresentativa = imagenes[miembros[0]];
    let nombre;
    try {
      nombre = await nombrarImagen(imagenRepresentativa);
    } catch (err) {
      console.warn(`[ordenar-excel] no se pudo nombrar el grupo ${g}, uso nombre generico: ${err.message}`);
      nombre = `grupo${g + 1}`;
    }
    miembros.forEach((indice) => {
      categoriaPorIndice[indice] = nombre;
    });
  }

  return categoriaPorIndice;
}

module.exports = { clasificarPorEmbeddings };
