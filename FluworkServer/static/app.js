document.addEventListener("DOMContentLoaded", () => {
  /* refs DOM ------------------------------------------------ */
  const $formBuscar = document.getElementById("formBuscar");
  const $cliente = document.getElementById("cliente");
  const $estado = document.getElementById("estado");
  const $tecnico = document.getElementById("tecnico");
  const $telefono = document.getElementById("telefono");
  const $email = document.getElementById("email");
  const $confirmacion = document.getElementById("confirmacion");
  const $modelo = document.getElementById("modelo");
  const $f_desde = document.getElementById("f_desde");
  const $f_hasta = document.getElementById("f_hasta");
  const $maxResultados = document.getElementById("maxResultados");

  const $tbody = document.getElementById("tblBody");
  const $totalTxt = document.getElementById("totalTxt");

  const $pagerTop = document.getElementById("pagerTop");
  const $btnFirstTop = document.getElementById("btnFirstTop");
  const $btnPrevTop = document.getElementById("btnPrevTop");
  const $btnNextTop = document.getElementById("btnNextTop");
  const $btnLastTop = document.getElementById("btnLastTop");
  const $pageInfoTop = document.getElementById("pageInfoTop");

  const $pager = document.getElementById("pager");
  const $btnFirst = document.getElementById("btnFirst");
  const $btnPrev = document.getElementById("btnPrev");
  const $btnNext = document.getElementById("btnNext");
  const $btnLast = document.getElementById("btnLast");
  const $pageInfo = document.getElementById("pageInfo");

  const $loadingRow = document.getElementById("loadingRow");
  const $noResultsRow = document.getElementById("noResults");

  $noResultsRow.classList.add("d-none");

  const PAGE_SIZE = 50;           // Fijo: resultados por página
  let currentOffset = 0;
  let currentTotal = 0;
  let lastPayload = {};

  function getMaxResultados() {
    return parseInt($maxResultados.value, 10) || 0;  // 0 = Todas
  }

  /* Eventos ------------------------------------------------ */
  $formBuscar.addEventListener("submit", (e) => {
    e.preventDefault();
    currentOffset = 0;
    doSearch();
  });

  [
    $btnPrev,
    $btnNext,
    $btnFirst,
    $btnLast,
    $btnPrevTop,
    $btnNextTop,
    $btnFirstTop,
    $btnLastTop,
  ].forEach((btn) => btn.addEventListener("click", onPager));

  $tbody.addEventListener("click", function (event) {
    const link = event.target.closest("a.open-ficha");
    const fila = event.target.closest("tr[data-ficha]");
    if (!link && !fila) return;

    const ficha = (link?.dataset.ficha || fila?.dataset.ficha || "").trim();
    if (!ficha) return;

    event.preventDefault();
    event.stopPropagation();

    window.location.href = `fluwork://${ficha}`;
  });

  /* ─────────────────────────────────────────────────────────
     FUNCIONES
     ───────────────────────────────────────────────────────── */

  function buildPayload() {
    return {
      problema: document.getElementById("problema")?.value.trim() || "",
      trabajos: document.getElementById("trabajos")?.value.trim() || "",
      serie: document.getElementById("serie")?.value.trim() || "",
      estado: $estado.value,
      cliente: $cliente.value.trim(),
      tecnico: $tecnico.value,
      telefono: $telefono.value.trim(),
      modelo: $modelo.value.trim(),
      email: $email?.value.trim() || "",
      confirmacion: $confirmacion?.value || "",
      fecha_desde: $f_desde.value,
      fecha_hasta: $f_hasta.value,
      limit: PAGE_SIZE,
      offset: currentOffset,
      max_total: getMaxResultados(),
    };
  }

  function doSearch() {
    lastPayload = buildPayload();

    $tbody.innerHTML = "";
    $noResultsRow.classList.add("d-none");
    $loadingRow.classList.remove("d-none");

    fetch("/api/buscar_avanzada", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(lastPayload),
    })
      .then(async (r) => {
        if (!r.ok) {
          const txt = await r.text().catch(() => "");
          throw new Error(`HTTP ${r.status} ${txt}`);
        }
        return r.json();
      })
      .then(renderTable)
      .catch((err) => {
        console.error("Error en búsqueda:", err);
        $loadingRow.classList.add("d-none");
        $tbody.innerHTML = `
          <tr><td colspan="7" class="text-center text-danger">
            Hubo un error al cargar los datos: ${err.message}
          </td></tr>`;
        $totalTxt.textContent = "Error al consultar el servidor.";
        $pagerTop.hidden = true;
        $pager.hidden = true;
      });
  }

  function renderTable(data) {
    $loadingRow.classList.add("d-none");

    if (!data || typeof data.total !== "number" || !Array.isArray(data.items)) {
      throw new Error("Respuesta inválida del backend");
    }

    currentTotal = data.total;
    $tbody.innerHTML = "";
    $noResultsRow.classList.add("d-none");

    if (data.total === 0) {
      $totalTxt.textContent = "No se encontraron resultados.";
    } else if (currentTotal <= PAGE_SIZE) {
      $totalTxt.textContent = `Total: ${data.total}`;
    } else {
      const page = Math.floor(currentOffset / PAGE_SIZE) + 1;
      const pages = Math.ceil(currentTotal / PAGE_SIZE);
      $totalTxt.textContent = `Total: ${data.total} (página ${page} de ${pages})`;
    }

    if (data.items.length === 0) {
      $noResultsRow.classList.remove("d-none");
      $tbody.appendChild($noResultsRow);
    } else {
      for (const r of data.items) {
        const tr = $tbody.insertRow();
        tr.dataset.ficha = r.ficha ?? "";

        tr.innerHTML = [
          `<td><a href="#" data-ficha="${r.ficha ?? ""}" class="link-primary open-ficha">${r.ficha ?? ""}</a></td>`,
          `<td>${r.fullname ?? ""}</td>`,
          `<td>${r.telefono ?? ""}</td>`,
          `<td>${r.modelo ?? ""}</td>`,
          `<td>${r.fecha ?? ""}</td>`,
          `<td>${r.estado ?? ""}</td>`,
          `<td>${r.tecnico ?? ""}</td>`,
        ].join("");
      }
    }

    const lastOffset = Math.max(0, Math.floor((currentTotal - 1) / PAGE_SIZE) * PAGE_SIZE);
    const currentPage = Math.floor(currentOffset / PAGE_SIZE) + 1;
    const totalPages = Math.ceil(currentTotal / PAGE_SIZE);
    const showPager = currentTotal > PAGE_SIZE;

    [$pagerTop, $pager].forEach((p) => { p.hidden = !showPager; });

    $btnPrevTop.disabled = currentOffset === 0;
    $btnFirstTop.disabled = currentOffset === 0;
    $btnNextTop.disabled = currentOffset >= lastOffset;
    $btnLastTop.disabled = currentOffset >= lastOffset;
    $pageInfoTop.textContent = `${currentPage}/${totalPages}`;

    $btnPrev.disabled = currentOffset === 0;
    $btnFirst.disabled = currentOffset === 0;
    $btnNext.disabled = currentOffset >= lastOffset;
    $btnLast.disabled = currentOffset >= lastOffset;
    $pageInfo.textContent = `${currentPage}/${totalPages}`;
  }

  function onPager(e) {
    const lastOffset = Math.max(0, Math.floor((currentTotal - 1) / PAGE_SIZE) * PAGE_SIZE);
    const id = e.currentTarget.id;

    if (id.includes("Prev"))
      currentOffset = Math.max(0, currentOffset - PAGE_SIZE);
    if (id.includes("Next"))
      currentOffset = Math.min(lastOffset, currentOffset + PAGE_SIZE);
    if (id.includes("First")) currentOffset = 0;
    if (id.includes("Last")) currentOffset = lastOffset;

    lastPayload.offset = currentOffset;

    $tbody.innerHTML = "";
    $noResultsRow.classList.add("d-none");
    $loadingRow.classList.remove("d-none");

    fetch("/api/buscar_avanzada", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(lastPayload),
    })
      .then(async (r) => {
        if (!r.ok) {
          const txt = await r.text().catch(() => "");
          throw new Error(`HTTP ${r.status} ${txt}`);
        }
        return r.json();
      })
      .then(renderTable)
      .catch((err) => {
        console.error("Error en paginación:", err);
        $loadingRow.classList.add("d-none");
        $tbody.innerHTML = `
          <tr><td colspan="7" class="text-center text-danger">
            Hubo un error al cargar los datos
          </td></tr>`;
        $totalTxt.textContent = "Error al consultar el servidor.";
        $pagerTop.hidden = true;
        $pager.hidden = true;
      });
  }
  /* Cargar versión del sistema */
  fetch("/api/version")
    .then((r) => r.json())
    .then((data) => {
      const el = document.getElementById("versionFooter");
      if (el && data.version) el.textContent = `Consulta Fluwork v${data.version}`;
    })
    .catch(() => {});
}); // fin DOMContentLoaded
