// Clustering jerarquico aglomerativo: empieza con cada vector como su propio grupo
// y va fusionando los dos grupos mas cercanos entre si, uno por vez, hasta que
// queda todo en un solo arbol (dendrograma). Sirve para dos cosas:
//   1) ordenarPorDendrograma: recorrer las hojas del arbol da un orden donde cada
//      item queda pegado a su vecino mas parecido en el espacio vectorial.
//   2) cortarEnGrupos: cortar el arbol a una distancia umbral da grupos discretos.

function distanciaCoseno(a, b) {
  let punto = 0;
  let normaA = 0;
  let normaB = 0;
  for (let i = 0; i < a.length; i++) {
    punto += a[i] * b[i];
    normaA += a[i] * a[i];
    normaB += b[i] * b[i];
  }
  if (normaA === 0 || normaB === 0) return 1;
  const similitud = punto / (Math.sqrt(normaA) * Math.sqrt(normaB));
  return 1 - similitud;
}

// Construye el dendrograma con "average linkage": la distancia entre dos grupos
// es el promedio de las distancias entre todos los pares de sus miembros.
function construirDendrograma(vectores) {
  const n = vectores.length;
  if (n === 0) return null;
  if (n === 1) return { tipo: "hoja", indice: 0, miembros: [0] };

  // cada nodo activo: {id, miembros: [indices originales], vectorPromedio}
  let nodos = vectores.map((v, i) => ({
    nodo: { tipo: "hoja", indice: i, miembros: [i] },
    miembros: [i],
  }));

  // matriz de distancias entre vectores originales (se reusa para promediar)
  const distOriginal = Array.from({ length: n }, () => new Array(n).fill(0));
  for (let i = 0; i < n; i++) {
    for (let j = i + 1; j < n; j++) {
      const d = distanciaCoseno(vectores[i], vectores[j]);
      distOriginal[i][j] = d;
      distOriginal[j][i] = d;
    }
  }

  function distanciaPromedioEntreGrupos(miembrosA, miembrosB) {
    let suma = 0;
    let cantidad = 0;
    for (const i of miembrosA) {
      for (const j of miembrosB) {
        suma += distOriginal[i][j];
        cantidad++;
      }
    }
    return suma / cantidad;
  }

  while (nodos.length > 1) {
    let mejorI = 0;
    let mejorJ = 1;
    let mejorDist = Infinity;

    for (let i = 0; i < nodos.length; i++) {
      for (let j = i + 1; j < nodos.length; j++) {
        const d = distanciaPromedioEntreGrupos(nodos[i].miembros, nodos[j].miembros);
        if (d < mejorDist) {
          mejorDist = d;
          mejorI = i;
          mejorJ = j;
        }
      }
    }

    const a = nodos[mejorI];
    const b = nodos[mejorJ];
    const nuevoNodo = {
      nodo: { tipo: "union", distancia: mejorDist, izquierda: a.nodo, derecha: b.nodo, miembros: [...a.miembros, ...b.miembros] },
      miembros: [...a.miembros, ...b.miembros],
    };

    nodos = nodos.filter((_, idx) => idx !== mejorI && idx !== mejorJ);
    nodos.push(nuevoNodo);
  }

  return nodos[0].nodo;
}

function ordenarPorDendrograma(dendrograma) {
  if (!dendrograma) return [];
  const orden = [];
  function recorrer(nodo) {
    if (nodo.tipo === "hoja") {
      orden.push(nodo.indice);
      return;
    }
    recorrer(nodo.izquierda);
    recorrer(nodo.derecha);
  }
  recorrer(dendrograma);
  return orden;
}

// Corta el arbol en grupos: cualquier union con distancia > umbral se considera
// un corte (sus dos lados quedan en grupos separados).
function cortarEnGrupos(dendrograma, umbral) {
  if (!dendrograma) return [];
  const grupos = [];

  function recolectarHojas(nodo, destino) {
    if (nodo.tipo === "hoja") {
      destino.push(nodo.indice);
      return;
    }
    recolectarHojas(nodo.izquierda, destino);
    recolectarHojas(nodo.derecha, destino);
  }

  function recorrer(nodo) {
    if (nodo.tipo === "hoja") {
      grupos.push([nodo.indice]);
      return;
    }
    if (nodo.distancia > umbral) {
      recorrer(nodo.izquierda);
      recorrer(nodo.derecha);
    } else {
      const miembros = [];
      recolectarHojas(nodo, miembros);
      grupos.push(miembros);
    }
  }

  recorrer(dendrograma);
  return grupos;
}

module.exports = { distanciaCoseno, construirDendrograma, ordenarPorDendrograma, cortarEnGrupos };
