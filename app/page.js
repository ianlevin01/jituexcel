"use client";

import { useState } from "react";

async function parsearRespuestaJson(respuesta) {
  const texto = await respuesta.text();
  try {
    return JSON.parse(texto);
  } catch {
    console.error("Respuesta no-JSON del servidor:", respuesta.status, texto.slice(0, 500));
    return {
      error: `El servidor respondió algo inesperado (status ${respuesta.status}): ${texto.slice(0, 200) || "(vacío)"}`,
    };
  }
}

async function subirArchivo(archivo, tipo) {
  const presignRes = await fetch("/api/excel-ab/upload-url", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ archivo: tipo }),
  });
  const presignData = await parsearRespuestaJson(presignRes);
  if (!presignRes.ok) throw new Error(presignData.error || "No se pudo iniciar la subida.");
  const { url, key } = presignData;

  const putRes = await fetch(url, {
    method: "PUT",
    headers: { "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
    body: archivo,
  });
  if (!putRes.ok) {
    const texto = await putRes.text().catch(() => "");
    throw new Error(`No se pudo subir el excel ${tipo.toUpperCase()} a S3 (status ${putRes.status}). ${texto.slice(0, 200)}`);
  }

  return key;
}

export default function Formulas() {
  const [archivo, setArchivo] = useState(null);
  const [procesando, setProcesando] = useState(false);
  const [mensaje, setMensaje] = useState(null);
  const [resumen, setResumen] = useState(null);

  const [archivoA, setArchivoA] = useState(null);
  const [archivoB, setArchivoB] = useState(null);
  const [procesandoAB, setProcesandoAB] = useState(false);
  const [mensajeAB, setMensajeAB] = useState(null);
  const [resumenAB, setResumenAB] = useState(null);

  async function procesar(e) {
    e.preventDefault();
    if (!archivo) return;

    setProcesando(true);
    setResumen(null);

    try {
      setMensaje({ tipo: "pendiente", texto: "Subiendo archivo..." });
      const presignRes = await fetch("/api/formulas/upload-url", { method: "POST" });
      const presignData = await parsearRespuestaJson(presignRes);
      if (!presignRes.ok) throw new Error(presignData.error || "No se pudo iniciar la subida.");
      const { url, key } = presignData;

      const putRes = await fetch(url, {
        method: "PUT",
        headers: { "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
        body: archivo,
      });
      if (!putRes.ok) {
        const texto = await putRes.text().catch(() => "");
        throw new Error(`No se pudo subir el archivo a S3 (status ${putRes.status}). ${texto.slice(0, 200)}`);
      }

      setMensaje({ tipo: "pendiente", texto: "Calculando fórmulas..." });
      const respuesta = await fetch("/api/formulas", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ key, tamanioEsperado: archivo.size }),
      });

      const data = await parsearRespuestaJson(respuesta);
      if (!respuesta.ok) throw new Error(data.error || `El servidor respondió con error (status ${respuesta.status}).`);

      const a = document.createElement("a");
      a.href = data.downloadUrl;
      a.download = "actualizado.xlsx";
      document.body.appendChild(a);
      a.click();
      a.remove();

      setResumen(data.resumen);
      setArchivo(null);
      setMensaje({ tipo: "ok", texto: "Listo, se descargó actualizado.xlsx." });
    } catch (err) {
      console.error("[formulas] error:", err);
      setMensaje({ tipo: "error", texto: err.message });
    } finally {
      setProcesando(false);
    }
  }

  async function procesarAB(e) {
    e.preventDefault();
    if (!archivoA || !archivoB) return;

    setProcesandoAB(true);
    setResumenAB(null);

    try {
      setMensajeAB({ tipo: "pendiente", texto: "Subiendo excel A..." });
      const keyA = await subirArchivo(archivoA, "a");

      setMensajeAB({ tipo: "pendiente", texto: "Subiendo excel B..." });
      const keyB = await subirArchivo(archivoB, "b");

      setMensajeAB({ tipo: "pendiente", texto: "Procesando..." });
      const respuesta = await fetch("/api/excel-ab", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          keyA,
          tamanioEsperadoA: archivoA.size,
          keyB,
          tamanioEsperadoB: archivoB.size,
        }),
      });

      const data = await parsearRespuestaJson(respuesta);
      if (!respuesta.ok) throw new Error(data.error || `El servidor respondió con error (status ${respuesta.status}).`);

      const a = document.createElement("a");
      a.href = data.downloadUrl;
      a.download = "actualizado.xlsx";
      document.body.appendChild(a);
      a.click();
      a.remove();

      setResumenAB(data.resumen);
      setArchivoA(null);
      setArchivoB(null);
      setMensajeAB({ tipo: "ok", texto: "Listo, se descargó actualizado.xlsx." });
    } catch (err) {
      console.error("[excel-ab] error:", err);
      setMensajeAB({ tipo: "error", texto: err.message });
    } finally {
      setProcesandoAB(false);
    }
  }

  return (
    <div className="page">
      <div>
        <nav className="top-nav">
          <a href="/">Fórmulas</a>
          <a href="/precios">Precios</a>
          <a href="/admin">Administración</a>
        </nav>

        <div className="card">
          <div className="card-header">
            <div className="icon">
              <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                <path d="M4 4h16v16H4z" stroke="white" strokeWidth="1.6" strokeLinejoin="round" />
                <path d="M4 9h16M9 4v16" stroke="white" strokeWidth="1.6" />
              </svg>
            </div>
            <h1>Actualizar fórmulas</h1>
          </div>
          <p className="subtitle">
            Subí el excel de productos y se agregan automáticamente las 8 columnas calculadas, igual que las
            cargás a mano.
          </p>

          <form onSubmit={procesar}>
            <div className="field">
              <label htmlFor="archivo">Excel de productos</label>
              <input
                type="file"
                id="archivo"
                accept=".xlsx"
                onChange={(e) => setArchivo(e.target.files[0] || null)}
              />
            </div>

            <button type="submit" disabled={!archivo || procesando}>
              Procesar y descargar
            </button>
          </form>

          {mensaje && <div className={`mensaje ${mensaje.tipo}`}>{mensaje.texto}</div>}

          {resumen && (
            <ul className="lista-clientes" style={{ marginTop: 12 }}>
              <li>
                <span className="cliente-nombre">Filas de producto</span>
                <span className="cliente-porcentaje">
                  {resumen.primeraFilaProducto}-{resumen.ultimaFilaProducto} ({resumen.cantidadFilasProducto})
                </span>
              </li>
              <li>
                <span className="cliente-nombre">
                  Amount of Goods ({resumen.columnaAmount}{resumen.filaAmount})
                </span>
                <span className="cliente-porcentaje">{resumen.valorJ}</span>
              </li>
            </ul>
          )}
        </div>

        <div className="card">
          <div className="card-header">
            <div className="icon">
              <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                <path d="M4 4h7v7H4zM13 4h7v7h-7zM4 13h7v7H4zM13 13h7v7h-7z" stroke="white" strokeWidth="1.6" strokeLinejoin="round" />
              </svg>
            </div>
            <h1>Combinar excel A y B</h1>
          </div>
          <p className="subtitle">
            Subí el excel A (el que se modifica) y el excel B (de donde se copia el precio). Por cada producto
            de A se busca su precio en B por código (columna B en A, columna A en B) y se pega en una columna
            nueva Q de A, corriendo todo lo que había de Q en adelante.
          </p>

          <form onSubmit={procesarAB}>
            <div className="field">
              <label htmlFor="archivoA">Excel A (a modificar)</label>
              <input
                type="file"
                id="archivoA"
                accept=".xlsx"
                onChange={(e) => setArchivoA(e.target.files[0] || null)}
              />
            </div>

            <div className="field">
              <label htmlFor="archivoB">Excel B (de donde se copia el precio)</label>
              <input
                type="file"
                id="archivoB"
                accept=".xlsx"
                onChange={(e) => setArchivoB(e.target.files[0] || null)}
              />
            </div>

            <button type="submit" disabled={!archivoA || !archivoB || procesandoAB}>
              Procesar y descargar
            </button>
          </form>

          {mensajeAB && <div className={`mensaje ${mensajeAB.tipo}`}>{mensajeAB.texto}</div>}

          {resumenAB && (
            <ul className="lista-clientes" style={{ marginTop: 12 }}>
              <li>
                <span className="cliente-nombre">Columna insertada en A</span>
                <span className="cliente-porcentaje">{resumenAB.columnaQ}</span>
              </li>
              <li>
                <span className="cliente-nombre">Precios encontrados y pegados</span>
                <span className="cliente-porcentaje">{resumenAB.cantidadValores}</span>
              </li>
              {resumenAB.cantidadSinPrecio > 0 && (
                <li>
                  <span className="cliente-nombre">Códigos de A sin precio en B</span>
                  <span className="cliente-porcentaje">{resumenAB.cantidadSinPrecio}</span>
                </li>
              )}
            </ul>
          )}
        </div>
      </div>
    </div>
  );
}
