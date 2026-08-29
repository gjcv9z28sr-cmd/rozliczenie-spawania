import JSZip from "jszip";

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
      return createJob(request, env);
    }

    if (request.method === "GET" && url.pathname === "/api/job") {
      return getJob(url.searchParams.get("id"), env);
    }

    return new Response("Not found", {
      status: 404
    });
  }
};

function parseOrderLine(line) {
  const columns = line
    .trim()
    .split(/\t+/)
    .map(value => value.trim());

  if (columns.length < 5) {
    throw new Error(
      "Wiersz musi zawierać minimum 5 kolumn oddzielonych tabulatorami."
    );
  }

  return {
    orderNumber: columns[0],
    date: columns[1],
    operator: columns[2],
    code: columns[3],
    description: columns.slice(4).join(" ").trim()
  };
}

function makeRootFolder(orderNumber, description) {
  return `${orderNumber.replace(/\//g, "_")} ${description}`;
}

async function createJob(request, env) {
  const form = await request.formData();

  const orderLine = String(
    form.get("orderLine") || ""
  );

  const archive = form.get("archive");

  if (!orderLine.trim()) {
    return json({
      error: "Brak danych zlecenia."
    }, 400);
  }

  if (!(archive instanceof File)) {
    return json({
      error: "Brak paczki ZIP."
    }, 400);
  }

  if (!archive.name.toLowerCase().endsWith(".zip")) {
    return json({
      error: "Paczka musi być plikiem ZIP."
    }, 400);
  }

  let order;

  try {
    order = parseOrderLine(orderLine);
  } catch (error) {
    return json({
      error: error.message
    }, 400);
  }

  const rootFolder = makeRootFolder(
    order.orderNumber,
    order.description
  );

  const id = crypto.randomUUID();

  /*
   * ---------------------------------------------------------
   * 1. ZAPISUJEMY ORYGINALNĄ PACZKĘ
   * ---------------------------------------------------------
   */

  const archiveName = archive.name.replace(
    /[^a-zA-Z0-9._-]/g,
    "_"
  );

  const inputKey =
    `jobs/${id}/input/${archiveName}`;

  await env.STORAGE.put(
    inputKey,
    archive.stream(),
    {
      httpMetadata: {
        contentType: "application/zip"
      },
      customMetadata: {
        orderNumber: order.orderNumber,
        date: order.date,
        operator: order.operator,
        code: order.code,
        description: order.description
      }
    }
  );

  /*
   * ---------------------------------------------------------
   * 2. POBIERAMY PUSTY SZABLON Z R2
   * ---------------------------------------------------------
   */

  const templateObject =
    await env.STORAGE.get(TEMPLATE_KEY);

  if (!templateObject) {
    return json({
      error:
        `Nie znaleziono szablonu ${TEMPLATE_KEY} w R2.`
    }, 500);
  }

  const templateBuffer =
    await templateObject.arrayBuffer();

  /*
   * ---------------------------------------------------------
   * 3. TWORZYMY KOPIĘ SZABLONU
   * ---------------------------------------------------------
   */

  const workbook =
    await JSZip.loadAsync(templateBuffer);

  await replaceB2InWorkbook(
    workbook,
    order.orderNumber
  );

  const resultBuffer =
    await workbook.generateAsync({
      type: "arraybuffer"
    });

  /*
   * ---------------------------------------------------------
   * 4. ZAPISUJEMY GOTOWY XLSX
   * ---------------------------------------------------------
   */

  const resultFileName =
    `${rootFolder}.xlsx`;

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

  /*
   * ---------------------------------------------------------
   * 5. TWORZYMY STRUKTURĘ KATALOGÓW
   * ---------------------------------------------------------
   */

  for (const folder of FOLDERS) {
    await env.STORAGE.put(
      `jobs/${id}/output/${rootFolder}/${folder}/`,
      ""
    );
  }

  /*
   * ---------------------------------------------------------
   * 6. INFORMACJA O ZADANIU
   * ---------------------------------------------------------
   */

  const job = {
    id,

    status: "RECEIVED",

    input: {
      rawLine: orderLine,
      orderNumber: order.orderNumber,
      date: order.date,
      operator: order.operator,
      code: order.code,
      description: order.description
    },

    output: {
      rootFolder,
      folders: FOLDERS,
      workbook: resultKey
    },

    inputKey,

    template: TEMPLATE_KEY,

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

  return json(job, 202);
}


/*
 * =========================================================
 * ZMIANA B2 W ARKUSZU ARKUSZ1
 * =========================================================
 *
 * Nie generujemy Excela od zera.
 * Modyfikujemy istniejący plik XLSX.
 */

async function replaceB2InWorkbook(
  zip,
  orderNumber
) {
  const workbookFile =
    zip.file("xl/workbook.xml");

  if (!workbookFile) {
    throw new Error(
      "Nieprawidłowy plik XLSX: brak xl/workbook.xml."
    );
  }

  const workbookXml =
    await workbookFile.async("string");

  const sheetMatch =
    workbookXml.match(
      /<sheet[^>]+name="Arkusz1"[^>]+r:id="([^"]+)"/
    );

  if (!sheetMatch) {
    throw new Error(
      "Nie znaleziono arkusza Arkusz1."
    );
  }

  const relationshipId =
    sheetMatch[1];

  const relsFile =
    zip.file(
      "xl/_rels/workbook.xml.rels"
    );

  if (!relsFile) {
    throw new Error(
      "Brak relacji workbook.xml.rels."
    );
  }

  const relsXml =
    await relsFile.async("string");

  const escapedId =
    relationshipId.replace(
      /[.*+?^${}()|[\]\\]/g,
      "\\$&"
    );

  const relationshipRegex =
    new RegExp(
      `<Relationship[^>]+Id="${escapedId}"[^>]+Target="([^"]+)"`
    );

  const relationshipMatch =
    relsXml.match(
      relationshipRegex
    );

  if (!relationshipMatch) {
    throw new Error(
      "Nie znaleziono pliku arkusza Arkusz1."
    );
  }

  let target =
    relationshipMatch[1];

  if (target.startsWith("/")) {
    target = target.substring(1);
  } else {
    target = "xl/" + target;
  }

  const sheetFile =
    zip.file(target);

  if (!sheetFile) {
    throw new Error(
      `Nie znaleziono pliku arkusza: ${target}`
    );
  }

  let sheetXml =
    await sheetFile.async("string");

  /*
   * B2 zapisujemy jako inline string.
   * Dzięki temu nie musimy modyfikować sharedStrings.xml.
   */

  const newCell =
    `<c r="B2" t="inlineStr"><is><t>${xmlEscape(orderNumber)}</t></is></c>`;

  const cellRegex =
    /<c\b[^>]*\br="B2"[^>]*>[\s\S]*?<\/c>/;

  if (cellRegex.test(sheetXml)) {
    sheetXml =
      sheetXml.replace(
        cellRegex,
        newCell
      );
  } else {
    const rowMatch =
      sheetXml.match(
        /<row\b[^>]*r="2"[^>]*>/
      );

    if (!rowMatch) {
      throw new Error(
        "Nie znaleziono wiersza 2 w Arkusz1."
      );
    }

    sheetXml =
      sheetXml.replace(
        rowMatch[0],
        `${rowMatch[0]}${newCell}`
      );
  }

  zip.file(
    target,
    sheetXml
  );
}

function xmlEscape(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

async function getJob(id, env) {
  if (!id) {
    return json({
      error: "Brak ID zadania."
    }, 400);
  }

  const object =
    await env.STORAGE.get(
      `jobs/${id}/job.json`
    );

  if (!object) {
    return json({
      error: "Nie znaleziono zadania."
    }, 404);
  }

  return new Response(
    await object.text(),
    {
      headers: {
        "content-type":
          "application/json; charset=utf-8"
      }
    }
  );
}

function json(value, status = 200) {
  return new Response(
    JSON.stringify(value, null, 2),
    {
      status,
      headers: {
        "content-type":
          "application/json; charset=utf-8"
      }
    }
  );
}

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
        "PACZKA PRZYJĘTA\\n\\n" +

        "Katalog:\\n" +

        data.output.rootFolder +

        "\\n\\n" +

        "Utworzono szablon rozliczenia:\\n" +

        data.output.workbook +

        "\\n\\n" +

        "ID zadania:\\n" +

        data.id;

    } catch (error) {

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
