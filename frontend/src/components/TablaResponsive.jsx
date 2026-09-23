import {
  useEffect,
  useMemo,
  useRef,
  useState
} from "react";

import * as XLSX from "xlsx";
import { saveAs } from "file-saver";

function TablaResponsive({ data = [] }) {

  const [pagina, setPagina] = useState(1);
  const [porPagina, setPorPagina] = useState(20);

  const [sortColumn, setSortColumn] = useState(null);
  const [sortDirection, setSortDirection] = useState("asc");

  const [columnasOcultas, setColumnasOcultas] = useState([]);

  const [busquedasFiltro, setBusquedasFiltro] = useState({});

  // Filtros por columna
  const [filtros, setFiltros] = useState({});

  // Filtro abierto actualmente
  const [filtroAbierto, setFiltroAbierto] = useState(null);

  const filtroRef = useRef(null);

  const dataSegura = Array.isArray(data)
    ? data
    : [];

  const columnasOriginales =
    dataSegura.length > 0
      ? Object.keys(dataSegura[0]).filter(
          (columna) =>
            !columna.startsWith("__")
        )
      : [];

  const columnas =
    columnasOriginales.filter(
      (col) => !columnasOcultas.includes(col)
    );

  // ==============================
  // CERRAR FILTRO AL HACER CLICK AFUERA
  // ==============================

  useEffect(() => {

    const cerrarFiltro = (event) => {

      if (
        filtroRef.current &&
        !filtroRef.current.contains(event.target)
      ) {
        setFiltroAbierto(null);
      }

    };

    document.addEventListener(
      "mousedown",
      cerrarFiltro
    );

    return () => {
      document.removeEventListener(
        "mousedown",
        cerrarFiltro
      );
    };

  }, []);

  // ==============================
  // VALORES ÚNICOS POR COLUMNA
  // ==============================

  const obtenerValoresUnicos = (columna) => {
    const valores = dataSegura.map(
      (fila) => String(fila[columna] ?? "")
    );

    const unicos = [...new Set(valores)]
      .sort((a, b) =>
        a.localeCompare(
          b,
          undefined,
          {
            numeric: true,
            sensitivity: "base"
          }
        )
      );

    const textoBusqueda =
      (busquedasFiltro[columna] || "")
        .toLowerCase()
        .trim();

    if (!textoBusqueda) {
      return unicos;
    }

    return unicos.filter((valor) =>
      valor
        .toLowerCase()
        .includes(textoBusqueda)
    );
  };

  // ==============================
  // FILTRADO
  // ==============================

  const dataFiltrada = useMemo(() => {

    return dataSegura.filter((fila) => {

      return Object.entries(filtros).every(
        ([columna, valoresSeleccionados]) => {

          if (
            !valoresSeleccionados ||
            valoresSeleccionados.length === 0
          ) {
            return true;
          }

          const valorFila =
            String(fila[columna] ?? "");

          return valoresSeleccionados.includes(
            valorFila
          );

        }
      );

    });

  }, [dataSegura, filtros]);

  // ==============================
  // ORDENAMIENTO
  // ==============================

  const dataOrdenada = useMemo(() => {

    let resultado = [...dataFiltrada];

    if (!sortColumn) {
      return resultado;
    }

    resultado.sort((a, b) => {

      const valorA =
        String(
          a[sortColumn] ?? ""
        ).toLowerCase();

      const valorB =
        String(
          b[sortColumn] ?? ""
        ).toLowerCase();

      return valorA.localeCompare(
        valorB,
        undefined,
        {
          numeric: true,
          sensitivity: "base"
        }
      ) *
        (sortDirection === "asc" ? 1 : -1);

    });

    return resultado;

  }, [
    dataFiltrada,
    sortColumn,
    sortDirection
  ]);

  // ==============================
  // PAGINACIÓN
  // ==============================

  const totalPaginas = Math.max(
    1,
    Math.ceil(
      dataOrdenada.length / porPagina
    )
  );

  useEffect(() => {
    setPagina(1);
  }, [
    filtros,
    porPagina,
    columnasOcultas
  ]);

  const inicio =
    (pagina - 1) * porPagina;

  const fin =
    inicio + porPagina;

  const datosPagina =
    dataOrdenada.slice(
      inicio,
      fin
    );

  // ==============================
  // ORDENAR
  // ==============================

  const ordenarColumna = (columna) => {

    if (sortColumn === columna) {

      setSortDirection(
        sortDirection === "asc"
          ? "desc"
          : "asc"
      );

    } else {

      setSortColumn(columna);
      setSortDirection("asc");

    }

  };

  // ==============================
  // FILTRAR VALORES
  // ==============================

  const toggleValorFiltro = (
    columna,
    valor
  ) => {

    const actuales =
      filtros[columna] || [];

    let nuevosValores;

    if (actuales.includes(valor)) {

      nuevosValores =
        actuales.filter(
          (item) => item !== valor
        );

    } else {

      nuevosValores = [
        ...actuales,
        valor
      ];

    }

    setFiltros({
      ...filtros,
      [columna]: nuevosValores
    });

  };

  const limpiarFiltroColumna = (
    columna
  ) => {

    const nuevosFiltros = {
      ...filtros
    };

    delete nuevosFiltros[columna];

    setFiltros(nuevosFiltros);

  };

  const limpiarTodosFiltros = () => {
    setFiltros({});
    setPagina(1);
  };

  // ==============================
  // EXPORTAR
  // ==============================

  const exportarExcel = () => {

    const dataExportada =
      dataOrdenada.map((fila) => {

        const nuevaFila = {};

        columnas.forEach((col) => {
          nuevaFila[col] = fila[col];
        });

        return nuevaFila;

      });

    const worksheet =
      XLSX.utils.json_to_sheet(
        dataExportada
      );

    const workbook =
      XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(
      workbook,
      worksheet,
      "Inventario"
    );

    const excelBuffer =
      XLSX.write(
        workbook,
        {
          bookType: "xlsx",
          type: "array"
        }
      );

    const archivo =
      new Blob(
        [excelBuffer],
        {
          type:
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        }
      );

    saveAs(
      archivo,
      "inventario.xlsx"
    );

  };

  const totalFiltrosActivos =
    Object.values(filtros)
      .filter(
        (valores) =>
          valores &&
          valores.length > 0
      )
      .length;

  return (

    <div
      className="
        bg-white
        rounded-3xl
        shadow-xl
        border
        border-slate-200
        overflow-hidden
        min-h-[700px]
        flex
        flex-col
      "
    >

      {/* CABECERA DE TABLA */}
      <div
        className="
          px-6
          py-4
          bg-gradient-to-r
          from-slate-800
          to-slate-900
          text-white
          flex
          flex-col
          lg:flex-row
          lg:justify-between
          lg:items-center
          gap-4
        "
      >

        <div>

          <h3 className="font-semibold text-lg">
            Inventario
          </h3>

          <p className="text-slate-300 text-sm">

            {dataOrdenada.length}
            {" "}
            de
            {" "}
            {dataSegura.length}
            {" "}
            registros

          </p>

        </div>

        <div
          className="
            flex
            flex-wrap
            gap-3
            items-center
          "
        >

          {/* MOSTRAR REGISTROS */}

          <div
            className="
              flex
              items-center
              gap-2
            "
          >

            <span className="text-sm">
              Mostrar:
            </span>

            <select
              value={porPagina}
              onChange={(e) =>
                setPorPagina(
                  Number(
                    e.target.value
                  )
                )
              }
              className="
                bg-white
                text-slate-800
                rounded-lg
                px-3
                py-2
              "
            >

              <option value={10}>
                10
              </option>

              <option value={20}>
                20
              </option>

              <option value={50}>
                50
              </option>

              <option value={100}>
                100
              </option>

            </select>

          </div>

          {/* LIMPIAR FILTROS */}

          {totalFiltrosActivos > 0 && (

            <button
              onClick={limpiarTodosFiltros}
              className="
                px-4
                py-2
                bg-slate-600
                text-white
                rounded-xl
                hover:bg-slate-500
                transition
              "
            >
              Limpiar filtros
              {" "}
              ({totalFiltrosActivos})
            </button>

          )}

          {/* EXPORTAR */}

          <button
            onClick={exportarExcel}
            className="
              px-4
              py-2
              bg-green-600
              text-white
              rounded-xl
              hover:bg-green-700
              transition
            "
          >
            📥 Exportar Excel
          </button>

        </div>

      </div>

      

      {/* TABLA */}
      <div
        className="
          overflow-auto
          flex-1
          min-h-[540px]
          max-h-[700px]
        "
      >

        <table
          className="
            min-w-[1200px]
            w-full
            text-sm
          "
        >

          <thead
            className="
              sticky
              top-0
              z-20
              bg-slate-700
              text-white
            "
          >

            <tr>

              {columnas.map(
                (columna) => {

                  const filtroActivo =
                    filtros[columna]?.length > 0;

                  return (

                    <th
                      key={columna}
                      className="
                        relative
                        p-0
                        text-left
                        whitespace-nowrap
                        font-semibold
                        border-b
                        border-slate-600
                      "
                    >

                      <div
                        className="
                          flex
                          items-center
                          justify-between
                          gap-3
                          px-4
                          py-3
                        "
                      >

                        {/* ORDENAR */}

                        <button
                          onClick={() =>
                            ordenarColumna(
                              columna
                            )
                          }
                          className="
                            flex
                            items-center
                            gap-2
                            hover:text-blue-200
                          "
                        >

                          <span>
                            {columna}
                          </span>

                          {sortColumn ===
                            columna && (

                            <span>

                              {sortDirection ===
                              "asc"
                                ? "▲"
                                : "▼"}

                            </span>

                          )}

                        </button>

                        {/* BOTÓN FILTRO */}

                        <button
                          onClick={(e) => {

                            e.stopPropagation();

                            setFiltroAbierto(
                              filtroAbierto ===
                                columna
                                ? null
                                : columna
                            );

                          }}
                          className={`
                            w-8
                            h-8
                            flex
                            items-center
                            justify-center
                            rounded-lg
                            transition

                            ${
                              filtroActivo
                                ? "bg-blue-500 text-white"
                                : "hover:bg-slate-600"
                            }
                          `}
                          title={`Filtrar ${columna}`}
                        >
                          ▾
                        </button>

                      </div>

                      {/* MENÚ FILTRO */}

                      {filtroAbierto ===
                        columna && (

                        <div
                          ref={filtroRef}
                          className={`
                            absolute
                            top-full
                            z-[100]
                            mt-2
                            w-72
                            max-w-[calc(100vw-40px)]
                            bg-white
                            text-slate-800
                            rounded-xl
                            shadow-2xl
                            border
                            border-slate-200

                            ${
                              columnas.indexOf(columna) <= 1
                                ? "left-0"
                                : "right-0"
                            }
                          `}
                        >

                          <div
                            className="
                              p-3
                              border-b
                              bg-slate-50
                            "
                          >

                            <p
                              className="
                                font-semibold
                                text-sm
                              "
                            >
                              Filtrar por {columna}
                            </p>

                            <input
                              type="text"
                              value={busquedasFiltro[columna] || ""}
                              onChange={(e) =>
                                setBusquedasFiltro({
                                  ...busquedasFiltro,
                                  [columna]: e.target.value
                                })
                              }
                              placeholder="Buscar valor..."
                              autoFocus
                              className="
                                mt-3
                                w-full
                                px-3
                                py-2
                                border
                                border-slate-300
                                rounded-lg
                                text-sm
                                text-slate-800
                                bg-white
                                outline-none
                                focus:ring-2
                                focus:ring-blue-500
                                focus:border-blue-500
                              "
                            />

                          </div>

                          {/* TODOS */}

                          <div
                            className="
                              px-3
                              py-2
                              border-b
                            "
                          >

                            <button
                              onClick={() =>
                                limpiarFiltroColumna(
                                  columna
                                )
                              }
                              className="
                                text-blue-600
                                text-sm
                                hover:underline
                              "
                            >
                              Mostrar todos
                            </button>

                          </div>

                          {/* VALORES */}

                          <div
                            className="
                              max-h-72
                              overflow-y-auto
                              p-2
                            "
                          >

                            {obtenerValoresUnicos(
                              columna
                            ).map(
                              (valor) => {

                                const seleccionado =
                                  filtros[
                                    columna
                                  ]?.includes(
                                    valor
                                  ) ||
                                  false;

                                return (

                                  <label
                                    key={valor}
                                    className="
                                      flex
                                      items-center
                                      gap-2
                                      px-2
                                      py-2
                                      rounded-lg
                                      hover:bg-slate-100
                                      cursor-pointer
                                    "
                                  >

                                    <input
                                      type="checkbox"
                                      checked={
                                        seleccionado
                                      }
                                      onChange={() =>
                                        toggleValorFiltro(
                                          columna,
                                          valor
                                        )
                                      }
                                    />

                                    <span
                                      className="
                                        truncate
                                        text-sm
                                      "
                                      title={
                                        valor ||
                                        "(Vacío)"
                                      }
                                    >

                                      {valor ||
                                        "(Vacío)"}

                                    </span>

                                  </label>

                                );

                              }
                            )}

                          </div>

                          <div
                            className="
                              p-3
                              border-t
                              bg-slate-50
                              flex
                              justify-end
                            "
                          >

                            <button
                              onClick={() =>
                                setFiltroAbierto(
                                  null
                                )
                              }
                              className="
                                px-4
                                py-2
                                bg-slate-800
                                text-white
                                rounded-lg
                                text-sm
                                hover:bg-slate-700
                              "
                            >
                              Cerrar
                            </button>

                          </div>

                        </div>

                      )}

                    </th>

                  );

                }
              )}

            </tr>

          </thead>

          <tbody>

            {datosPagina.length > 0 ? (

              datosPagina.map(
                (fila, index) => (

                  <tr
                    key={`${pagina}-${index}`}
                    className={`
                      ${
                        index % 2 === 0
                          ? "bg-white"
                          : "bg-slate-50"
                      }

                      hover:bg-blue-50
                      transition
                    `}
                  >

                    {columnas.map(
                      (columna) => (

                        <td
                          key={columna}
                          className="
                            p-3
                            border-b
                            border-slate-100
                            whitespace-nowrap
                          "
                        >

                          {String(
                            fila[
                              columna
                            ] ?? ""
                          )}

                        </td>

                      )
                    )}

                  </tr>

                )
              )

            ) : (

              <tr>

                <td
                  colSpan={
                    columnas.length || 1
                  }
                  className="
                    text-center
                    p-10
                    text-slate-500
                  "
                >
                  No se encontraron registros
                  con los filtros seleccionados.
                </td>

              </tr>

            )}

          </tbody>

        </table>

      </div>

      {/* FOOTER */}
      <div
        className="
          p-4
          bg-slate-50
          border-t
          flex
          flex-col
          md:flex-row
          md:justify-between
          md:items-center
          gap-4
        "
      >

        <div
          className="
            text-sm
            text-slate-600
          "
        >

          Mostrando{" "}

          <strong>
            {dataOrdenada.length === 0
              ? 0
              : inicio + 1}
          </strong>

          {" "}a{" "}

          <strong>
            {Math.min(
              fin,
              dataOrdenada.length
            )}
          </strong>

          {" "}de{" "}

          <strong>
            {dataOrdenada.length}
          </strong>

          {" "}registros

        </div>

        <div
          className="
            flex
            items-center
            gap-2
          "
        >

          <button
            disabled={pagina === 1}
            onClick={() =>
              setPagina(
                (prev) => prev - 1
              )
            }
            className="
              px-4
              py-2
              rounded-xl
              bg-slate-200
              hover:bg-slate-300
              disabled:opacity-40
              disabled:cursor-not-allowed
            "
          >
            ← Anterior
          </button>

          <div
            className="
              px-4
              py-2
              rounded-xl
              bg-blue-600
              text-white
              font-semibold
            "
          >

            {pagina}
            {" / "}
            {totalPaginas}

          </div>

          <button
            disabled={
              pagina >=
              totalPaginas
            }
            onClick={() =>
              setPagina(
                (prev) => prev + 1
              )
            }
            className="
              px-4
              py-2
              rounded-xl
              bg-slate-200
              hover:bg-slate-300
              disabled:opacity-40
              disabled:cursor-not-allowed
            "
          >
            Siguiente →
          </button>

        </div>

      </div>

    </div>

  );
}

export default TablaResponsive;