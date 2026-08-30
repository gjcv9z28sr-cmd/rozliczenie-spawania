import JSZip from "jszip";
import * as XLSX from "xlsx";

const TEMPLATE_KEY = "#XXXXX.xlsx";

const FOLDERS = [
  "0) Materiały",
  "1) PW OPL",
  "2) PW OBIEKT",
  "3) DP OPL",
  "4) DP PEŁNA",
  "5) Spawanie",
  "6) Rozliczenie"
];

export default {
  async fetch(request, env) {

    const url = new URL(request.url);

    if (request.method === "GET" && url.pathname === "/") {
      return new Response(page(), {
        headers: {
          "content-type": "text/html; charset=utf-8"
        }
      });
    }

    if (request.method === "POST" && url.pathname === "/api/job") {
      return processJob(request, env);
    }

    return new Response("Not found", {
      status: 404
    });
  }
};


/* =========================================================
   GŁÓWNY PROCES
   ========================================================= */

async function processJob(request, env) {

  const form = await request.formData();

  const orderLine =
    String(form.get("orderLine") || "").trim();

  const uploadedZip =
    form.get("archive");

  if (!orderLine) {
    return json({
      error: "Brak wiersza danych zlecenia."
    }, 400);
  }

  if (!(uploadedZip instanceof File)) {
    return json({
      error: "Brak paczki ZIP."
    }, 400);
  }

  if (
    !uploadedZip.name
      .toLowerCase()
      .endsWith(".zip")
  ) {
    return json({
      error: "Przesłany plik nie jest ZIP."
    }, 400);
  }


  /* -------------------------------------------------------
     DANE ZLECENIA
     ------------------------------------------------------- */

  const order = parseOrderLine(orderLine);


  /* -------------------------------------------------------
     ID ZADANIA
     ------------------------------------------------------- */

  const id = crypto.randomUUID();


  /* -------------------------------------------------------
     ODCZYT ZIP
     ------------------------------------------------------- */

  const zipBuffer =
    await uploadedZip.arrayBuffer();

  const inputZip =
    await JSZip.loadAsync(zipBuffer);


  /* -------------------------------------------------------
     SZUKAMY XLSX ZACZYNAJĄCEGO SIĘ OD #
     ------------------------------------------------------- */

  const sourceXlsx =
    await findSourceXlsx(inputZip);

  if (!sourceXlsx) {
    return json({
      error:
        "W paczce nie znaleziono pliku XLSX zaczynającego się od #."
    }, 400);
  }


  /* -------------------------------------------------------
     ODCZYTUJEMY XLSX ŹRÓDŁOWY
     ------------------------------------------------------- */

  const sourceBuffer =
    await sourceXlsx.async("arraybuffer");

  const sourceWorkbook =
    XLSX.read(sourceBuffer, {
      type: "array",
      cellDates: true
    });


  /* -------------------------------------------------------
     NUMER ROZLICZENIA
     ------------------------------------------------------- */

  const runNumber =
    extractRunNumber(
      sourceXlsx.name
    );

  if (!runNumber) {
    return json({
      error:
        "Nie udało się odczytać numeru rozpoczynającego się od #."
    }, 400);
  }


  /* -------------------------------------------------------
     INFORMACJE O PRZEBIEGACH
     ------------------------------------------------------- */

  const runInfo =
    extractRunInformation(
      sourceWorkbook
    );


  /* -------------------------------------------------------
     SZABLON MACIERZYSTY
     ------------------------------------------------------- */

  const templateObject =
    await env.STORAGE.get(TEMPLATE_KEY);

  if (!templateObject) {
    return json({
      error:
        `Nie znaleziono ${TEMPLATE_KEY} w R2.`
    }, 500);
  }

  const templateBuffer =
    await templateObject.arrayBuffer();


  /* -------------------------------------------------------
     OTWIERAMY SZABLON
     ------------------------------------------------------- */

  const workbook =
    XLSX.read(templateBuffer, {
      type: "array",
      cellStyles: true,
      cellDates: true
    });


  /* -------------------------------------------------------
     UZUPEŁNIAMY ROZLICZENIE
     ------------------------------------------------------- */

  fillSettlement(
    workbook,
    runNumber,
    runInfo
  );


  /* -------------------------------------------------------
     GENERUJEMY XLSX
     ------------------------------------------------------- */

  const resultBuffer =
    XLSX.write(workbook, {
      bookType: "xlsx",
      type: "array",
      cellStyles: true
    });


  /* -------------------------------------------------------
     NAZWA WYNIKU
     ------------------------------------------------------- */

  const resultFileName =
    `${runNumber}.xlsx`;


  /* -------------------------------------------------------
     NAZWA KATALOGU
     ------------------------------------------------------- */

  const rootFolder =
    makeRootFolder(
      order.orderNumber,
      order.description
    );


  /* -------------------------------------------------------
     ZAPIS PACZKI ŹRÓDŁOWEJ
     ------------------------------------------------------- */

  const inputKey =
    `jobs/${id}/input/${uploadedZip.name}`;

  await env.STORAGE.put(
    inputKey,
    zipBuffer,
    {
      httpMetadata: {
        contentType: "application/zip"
      }
    }
  );


  /* -------------------------------------------------------
     TWORZYMY KATALOGI
     ------------------------------------------------------- */

  for (const folder of FOLDERS) {

    await env.STORAGE.put(
      `jobs/${id}/output/${rootFolder}/${folder}/`,
      ""
    );

  }


  /* -------------------------------------------------------
     ZAPIS ROZLICZENIA
     ------------------------------------------------------- */

  const resultKey =
    `jobs/${id}/output/${rootFolder}/6) Rozliczenie/${resultFileName}`;

  await env.STORAGE.put(
    resultKey,
    resultBuffer,
    {
      httpMetadata: {
        contentType:
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      }
    }
  );


  /* -------------------------------------------------------
     INFORMACJA O WYNIKU
     ------------------------------------------------------- */

  const job = {

    id,

    status: "COMPLETED",

    input: {
      rawLine: orderLine,
      ...order
    },

    source: {
      file: sourceXlsx.name,
      runNumber
    },

    extracted: runInfo,

    output: {
      rootFolder,
      workbook: resultKey
    },

    createdAt:
      new Date().toISOString()

  };


  await env.STORAGE.put(
    `jobs/${id}/job.json`,
    JSON.stringify(job, null, 2),
    {
      httpMetadata: {
        contentType:
          "application/json"
      }
    }
  );


  return json(job, 200);
}


/* =========================================================
   PARSOWANIE WIERSZA ZLECENIA
   ========================================================= */

function parseOrderLine(line) {

  const columns =
    line
      .split(/\t+/)
      .map(x => x.trim())
      .filter(Boolean);

  if (columns.length < 5) {
    throw new Error(
      "Wiersz zlecenia musi zawierać minimum 5 kolumn."
    );
  }

  return {

    orderNumber: columns[0],

    date: columns[1],

    operator: columns[2],

    code: columns[3],

    description:
      columns.slice(4).join(" ")

  };
}


/* =========================================================
   SZUKANIE XLSX Z #
   ========================================================= */

async function findSourceXlsx(zip) {

  let found = null;

  zip.forEach((path, file) => {

    if (found) return;

    const name =
      path.split("/").pop();

    if (
      !file.dir &&
      name.startsWith("#") &&
      name.toLowerCase().endsWith(".xlsx")
    ) {
      found = file;
      found.name = name;
    }

  });

  return found;
}


/* =========================================================
   NUMER PRZEBIEGU
   ========================================================= */

function extractRunNumber(filename) {

  const match =
    filename.match(/^#[^.\s]+/);

  if (!match) {
    return null;
  }

  return match[0];
}


/* =========================================================
   ODCZYT INFORMACJI O PRZEBIEGACH
   ========================================================= */

function extractRunInformation(workbook) {

  const result = {
    splicedMuffs: []
  };


  for (const sheetName of workbook.SheetNames) {

    const sheet =
      workbook.Sheets[sheetName];

    const rows =
      XLSX.utils.sheet_to_json(
        sheet,
        {
          header: 1,
          defval: ""
        }
      );


    /*
     * Szukamy sekcji "Informacje o przebiegach".
     */

    let sectionFound = false;


    for (
      let rowIndex = 0;
      rowIndex < rows.length;
      rowIndex++
    ) {

      const row =
        rows[rowIndex];

      const text =
        row
          .map(value =>
            String(value)
              .trim()
              .toLowerCase()
          )
          .join(" ");


      if (
        text.includes(
          "informacje o przebiegach"
        )
      ) {

        sectionFound = true;

        /*
         * Po znalezieniu sekcji analizujemy
         * kolejne wiersze.
         */

        for (
          let i = rowIndex + 1;
          i < rows.length;
          i++
        ) {

          const data =
            rows[i];

          const joined =
            data
              .map(value =>
                String(value).trim()
              )
              .join(" | ");


          const lower =
            joined.toLowerCase();


          if (
            lower.includes(
              "nie istnieje - przecina tube"
            )
          ) {

            const name =
              findMuffName(data);

            const splices = 12;

            if (name) {

              result.splicedMuffs.push({

                name,

                splices,

                tubeCut: true,

                prepareCableEnd: 1

              });

            }

          }

          else if (
            lower.includes(
              "nie istnieje"
            )
          ) {

            const name =
              findMuffName(data);

            const splices =
              findSpliceCount(
                data,
                rows,
                i
              );

            if (name) {

              result.splicedMuffs.push({

                name,

                splices,

                tubeCut: false,

                prepareCableEnd: 0

              });

            }

          }

        }

        break;
      }
    }


    if (!sectionFound) {
      throw new Error(
        "Nie znaleziono sekcji 'Informacje o przebiegach'."
      );
    }

  }


  return result;
}


/* =========================================================
   NAZWA MUFY
   ========================================================= */

function findMuffName(row) {

  /*
   * Szukamy pierwszej sensownej wartości
   * zawierającej typowy identyfikator mufy.
   */

  for (const value of row) {

    const text =
      String(value).trim();

    if (!text) continue;

    const lower =
      text.toLowerCase();

    if (
      lower.includes("nie istnieje")
    ) {
      continue;
    }

    if (
      lower.includes("spaw")
    ) {
      continue;
    }

    if (
      lower.includes("liczba")
    ) {
      continue;
    }

    /*
     * Pomijamy oczywiste nagłówki.
     */

    if (
      lower ===
      "informacje o przebiegach"
    ) {
      continue;
    }

    return text;
  }

  return null;
}


/* =========================================================
   LICZBA SPAWÓW
   ========================================================= */

function findSpliceCount(
  row,
  rows,
  rowIndex
) {

  /*
   * Najpierw szukamy liczby w tym samym wierszu.
   */

  for (let i = 0; i < row.length; i++) {

    const value =
      String(row[i]).trim();

    if (
      /liczba spawów pomiędzy portem a włóknem/i
        .test(value)
    ) {

      const next =
        row[i + 1];

      const number =
        Number(next);

      if (
        Number.isFinite(number)
      ) {
        return number;
      }

    }

  }


  /*
   * Następnie sprawdzamy sąsiednie
   * wiersze.
   */

  for (
    let offset = 1;
    offset <= 5;
    offset++
  ) {

    const index =
      rowIndex + offset;

    if (
      index >= rows.length
    ) {
      break;
    }

    const candidate =
      rows[index];

    const text =
      candidate
        .map(x => String(x))
        .join(" ");


    const match =
      text.match(
        /liczba spawów pomiędzy portem a włóknem[^0-9]*([0-9]+)/i
      );

    if (match) {
      return Number(match[1]);
    }

  }


  /*
   * Jeżeli informacja istnieje, ale parser
   * jej nie znalazł, nie zgadujemy.
   */

  return 0;
}


/* =========================================================
   UZUPEŁNIENIE SZABLONU
   ========================================================= */

function fillSettlement(
  workbook,
  runNumber,
  runInfo
) {

  const sheet =
    workbook.Sheets["Arkusz1"];

  if (!sheet) {
    throw new Error(
      "W szablonie nie znaleziono arkusza Arkusz1."
    );
  }


  /*
   * B2 = numer przebiegu
   */

  sheet["B2"] = {
    t: "s",
    v: runNumber
  };


  /*
   * Mufy zaczynamy od F2.
   *
   * F = 6
   */

  const startColumn = 6;


  for (
    let i = 0;
    i < runInfo.splicedMuffs.length;
    i++
  ) {

    const muff =
      runInfo.splicedMuffs[i];

    const column =
      startColumn + i;


    /*
     * Wiersz 2 — nazwa mufy
     */

    const nameCell =
      XLSX.utils.encode_cell({
        r: 1,
        c: column - 1
      });

    sheet[nameCell] = {
      t: "s",
      v: muff.name
    };


    /*
     * Wiersz 4 — liczba spawów
     */

    const spliceCell =
      XLSX.utils.encode_cell({
        r: 3,
        c: column - 1
      });

    sheet[spliceCell] = {
      t: "n",
      v: muff.splices
    };


    /*
     * Przygotowanie końcówki kabla
     *
     * W szablonie jest to wiersz 4
     * prac — Excelowy numer wiersza 5,
     * czyli indeks 4.
     *
     * Ponieważ nie chcemy jeszcze zgadywać
     * kolumny/pozycji dodatkowej pracy,
     * zapisujemy ją w modelu wyniku.
     */

  }


  /*
   * Aktualizacja zakresu arkusza.
   */

  const range =
    XLSX.utils.decode_range(
      sheet["!ref"] || "A1"
    );

  const requiredLastColumn =
    startColumn +
    runInfo.splicedMuffs.length -
    1;

  if (
    requiredLastColumn >
    range.e.c
  ) {

    range.e.c =
      requiredLastColumn;

    sheet["!ref"] =
      XLSX.utils.encode_range(range);

  }

}


/* =========================================================
   NAZWA KATALOGU
   ========================================================= */

function makeRootFolder(
  orderNumber,
  description
) {

  return (
    orderNumber.replace(/\//g, "_") +
    " " +
    description
  );

}


/* =========================================================
   JSON
   ========================================================= */

function json(
  value,
  status = 200
) {

  return new Response(
    JSON.stringify(
      value,
      null,
      2
    ),
    {
      status,

      headers: {
        "content-type":
          "application/json; charset=utf-8"
      }
    }
  );

}


/* =========================================================
   FORMULARZ
   ========================================================= */

function page() {

  return `<!doctype html>

<html lang="pl">

<head>

<meta charset="utf-8">

<meta
  name="viewport"
  content="width=device-width,initial-scale=1"
>

<title>Agent2</title>

<style>

body {

  font-family:
    -apple-system,
    BlinkMacSystemFont,
    sans-serif;

  max-width: 700px;

  margin: auto;

  padding: 24px;

}

textarea,
input,
button {

  width: 100%;

  box-sizing: border-box;

  padding: 14px;

  margin-top: 8px;

  border-radius: 10px;

  border: 1px solid #aaa;

  font-size: 16px;

}

textarea {

  min-height: 130px;

}

button {

  margin-top: 20px;

  font-weight: 700;

}

#status {

  margin-top: 24px;

  white-space: pre-wrap;

  word-break: break-word;

}

</style>

</head>

<body>

<h1>Agent2</h1>

<form id="jobForm">

<label>

Cały wiersz danych zlecenia

<textarea
name="orderLine"
placeholder="64/W/2025&#9;03.06.2025&#9;PLAY&#9;KWOT2e&#9;Zlecenie uzgodnienia oraz budowa kabla do BTSa WAR1439A Otwock Tadeusza 22 72j"
required
></textarea>

</label>

<label>

Paczka ZIP

<input
name="archive"
type="file"
accept=".zip,application/zip"
required
>

</label>

<button type="submit">

PRZYJMIJ ZLECENIE

</button>

</form>

<div id="status"></div>

<script>

const form =
  document.getElementById("jobForm");

const status =
  document.getElementById("status");

form.addEventListener(
  "submit",
  async event => {

    event.preventDefault();

    status.textContent =
      "Przetwarzanie...";

    try {

      const response =
        await fetch(
          "/api/job",
          {
            method: "POST",
            body:
              new FormData(form)
          }
        );

      const data =
        await response.json();

      if (!response.ok) {

        throw new Error(
          data.error ||
          "Wystąpił błąd."
        );

      }

      status.textContent =
        "ROZLICZENIE UTWORZONE\\n\\n" +

        "Numer przebiegu: " +
        data.source.runNumber +

        "\\n\\nKatalog:\\n" +
        data.output.rootFolder +

        "\\n\\nPlik:\\n" +
        data.output.workbook +

        "\\n\\nID:\\n" +
        data.id;

    }

    catch (error) {

      status.textContent =
        "BŁĄD:\\n" +
        error.message;

    }

  }
);

</script>

</body>

</html>`;
}
