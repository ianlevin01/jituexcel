"use client";

import { useEffect, useMemo, useState } from "react";

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

export default function Precios() {
  const [clientes, setClientes] = useState([]);
  const [busqueda, setBusqueda] = useState("");
  const [clienteSeleccionado, setClienteSeleccionado] = useState(null);
  const [archivo, setArchivo] = useState(null);
  const [procesando, setProcesando] = useState(false);
  const [mensaje, setMensaje] = useState(null);

  const [archivoA, setArchivoA] = useState(null);
  const [archivoB, setArchivoB] = useState(null);
  const [procesandoAB, setProcesandoAB] = useState(false);
  const [mensajeAB, setMensajeAB] = useState(null);
  const [resumenAB, setResumenAB] = useState(null);

  useEffect(() => {
    fetch("/api/clientes")
      .then((r) => r.json())
      .then(setClientes)
      .catch(() => setMensaje({ tipo: "error", texto: "No se pudo cargar la lista de clientes." }));
  }, []);

  const clientesFiltrados = useMemo(() => {
    const q = busqueda.trim().toLowerCase();
    if (!q) return clientes;
    return clientes.filter((c) => c.nombre.toLowerCase().includes(q));
  }, [clientes, busqueda]);

  function cambiarBusqueda(valor) {
    setBusqueda(valor);
    if (clienteSeleccionado && valor !== clienteSeleccionado.nombre) {
      setClienteSeleccionado(null);
    }
  }

  function limpiarSeleccion() {
    setClienteSeleccionado(null);
    setBusqueda("");
  }

  async function descargar() {
    if (!clienteSeleccionado || !archivo) return;

    setProcesando(true);

    try {
      setMensaje({ tipo: "pendiente", texto: "Subiendo archivo..." });
      const presignRes = await fetch("/api/reprice/upload-url", { method: "POST" });
      const presignData = await presignRes.json().catch(() => ({}));
      if (!presignRes.ok) throw new Error(presignData.error || "No se pudo iniciar la subida.");
      const { url: uploadUrl, key } = presignData;

      const putRes = await fetch(uploadUrl, {
        method: "PUT",
        headers: { "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
        body: archivo,
      });
      if (!putRes.ok) throw new Error("No se pudo subir el archivo a S3.");

      setMensaje({ tipo: "pendiente", texto: "Procesando..." });
      const respuesta = await fetch("/api/reprice", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ clienteId: clienteSeleccionado.clienteId, key, tamanioEsperado: archivo.size }),
      });

      const data = await respuesta.json().catch(() => ({}));
      if (!respuesta.ok) throw new Error(data.error || "No se pudo generar el excel.");

      const a = document.createElement("a");
      a.href = data.downloadUrl;
      a.download = "actualizados.xlsx";
      document.body.appendChild(a);
      a.click();
      a.remove();

      const filasOmitidas = Number(data.filasOmitidas || 0);
      let texto = `Listo. Se descargó actualizados.xlsx para "${clienteSeleccionado.nombre}" (${clienteSeleccionado.porcentaje}%).`;
      if (filasOmitidas > 0) {
        texto += `\nSe omitieron ${filasOmitidas} fila(s) sin coincidencia en el excel base.`;
      }
      setMensaje({ tipo: "ok", texto });
    } catch (err) {
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
                <path d="M4 4h11l5 5v11a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1V5a1 1 0 0 1 1-1Z" stroke="white" strokeWidth="1.6" strokeLinejoin="round" />
                <path d="M14 4v5h5" stroke="white" strokeWidth="1.6" strokeLinejoin="round" />
                <path d="M8 13.5h2M8 16.5h5M14.5 13.5v3" stroke="white" strokeWidth="1.6" strokeLinecap="round" />
              </svg>
            </div>
            <h1>Descargar precios de cliente</h1>
          </div>
          <p className="subtitle">
            Buscá un cliente, subí su excel y descargá los precios actualizados con su porcentaje.
          </p>

          <div className="field">
            <label htmlFor="busqueda">Cliente</label>
            <input
              type="text"
              id="busqueda"
              placeholder="Buscar por nombre..."
              value={busqueda}
              onChange={(e) => cambiarBusqueda(e.target.value)}
              autoComplete="off"
            />
          </div>

          {busqueda.trim() && !clienteSeleccionado && (
            <ul className="lista-clientes" style={{ marginBottom: 20 }}>
              {clientesFiltrados.length === 0 && <li>Sin resultados.</li>}
              {clientesFiltrados.map((c) => (
                <li
                  key={c.clienteId}
                  className="seleccionable"
                  onClick={() => {
                    setClienteSeleccionado(c);
                    setBusqueda(c.nombre);
                  }}
                >
                  <span className="cliente-nombre">{c.nombre}</span>
                  <span className="cliente-porcentaje">{c.porcentaje}%</span>
                </li>
              ))}
            </ul>
          )}

          {clienteSeleccionado && (
            <ul className="lista-clientes" style={{ marginBottom: 20 }}>
              <li className="seleccionado">
                <span className="cliente-nombre">{clienteSeleccionado.nombre}</span>
                <span className="cliente-porcentaje">{clienteSeleccionado.porcentaje}%</span>
                <div className="acciones-fila">
                  <button type="button" className="secundario" onClick={limpiarSeleccion}>
                    Cambiar
                  </button>
                </div>
              </li>
            </ul>
          )}

          <div className="field">
            <label htmlFor="archivo">Excel a modificar</label>
            <input
              type="file"
              id="archivo"
              accept=".xlsx"
              onChange={(e) => setArchivo(e.target.files[0] || null)}
            />
          </div>

          <button onClick={descargar} disabled={!clienteSeleccionado || !archivo || procesando}>
            Descargar
          </button>

          {mensaje && <div className={`mensaje ${mensaje.tipo}`}>{mensaje.texto}</div>}
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
            Subí el excel A (el que se modifica) y el excel B (de donde se copia el precio). Se agregan 3
            columnas nuevas en A con el precio de B y los cálculos derivados.
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
                <span className="cliente-nombre">Columnas agregadas</span>
                <span className="cliente-porcentaje">
                  {resumenAB.columnaQ}, {resumenAB.columnaR}, {resumenAB.columnaS}
                </span>
              </li>
              <li>
                <span className="cliente-nombre">Filas procesadas</span>
                <span className="cliente-porcentaje">{resumenAB.cantidadFilasProcesadas}</span>
              </li>
              {resumenAB.filasOmitidas > 0 && (
                <li>
                  <span className="cliente-nombre">Omitidas (sin precio en B)</span>
                  <span className="cliente-porcentaje">{resumenAB.filasOmitidas}</span>
                </li>
              )}
            </ul>
          )}
        </div>
      </div>
    </div>
  );
}
