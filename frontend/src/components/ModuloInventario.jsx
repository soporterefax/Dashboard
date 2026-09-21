import TablaResponsive from "./TablaResponsive";

function ModuloInventario({
  titulo,
  data = []
}) {

  const dataSegura = Array.isArray(data)
    ? data
    : [];

  return (
    <div
      className="
        max-w-[1800px]
        mx-auto
        space-y-6
      "
    >

      {/* HEADER SIMPLE */}
      <div
        className="
          bg-gradient-to-r
          from-slate-900
          via-slate-800
          to-slate-900
          text-white
          p-8
          rounded-3xl
          shadow-xl
        "
      >
        <h1 className="text-4xl font-bold">
          {titulo}
        </h1>

        <p className="text-slate-300 mt-2">
          Gestión y análisis de activos TI
        </p>
      </div>

      {/* TABLA */}
      <TablaResponsive
        data={dataSegura}
      />

    </div>
  );
}

export default ModuloInventario;