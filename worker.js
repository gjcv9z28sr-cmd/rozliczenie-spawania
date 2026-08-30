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

    if (
      request.method === "GET" &&
      url.pathname === "/"
    ) {
      return new Response(page(), {
        headers: {
          "content-type":
            "text/html; charset=utf-8"
        }
      });
    }

    if (
      request.method === "POST" &&
      url.pathname === "/api/job"
    ) {
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
    String(
      form.get("orderLine") || ""
    ).trim();

  const uploadedZip =
    form.get("archive");

  if (!orderLine) {
    return json(
      {
        error:
          "Brak wiersza danych zlecenia."
      },
      400
    );
  }

  if (!(uploadedZip instanceof File)) {
    return json(
      {
        error:
          "Brak paczki ZIP."
      },
      400
    );
  }

  if (
    !uploadedZip.name
      .toLowerCase()
      .endsWith(".zip")
  ) {
    return json(
      {
        error:
          "Przesłany plik nie jest ZIP."
      },
      400
    );
  }


  /* -------------------------------------------------------
     DANE ZLECENIA
     ------------------------------------------------------- */

  const order =
    parseOrderLine(orderLine);


  /* -------------------------------------------------------
     ZIP
     ------------------------------------------------------- */

  const zipBuffer =
    await uploadedZip.arrayBuffer();

  const zip =
    await JSZip.loadAsync(zipBuffer);


  /* -------------------------------------------------------
     SZUKAMY XLSX ZACZYNAJĄCEGO SIĘ OD #
     ------------------------------------------------------- */

  const source =
    await findSourceXlsx(zip);

  if (!source) {
    return json(
      {
        error:
          "W paczce nie znaleziono pliku XLSX zaczynającego się od #."
      },
      400
    );
  }


  /*
   * NUMER MUSI BYĆ DOKŁADNIE:
   *
   * # + 5 cyfr
   *
   * np. #21482
   */

  const runNumber =
    extractFiveDigitNumber(
      source.name
    );

  if (!runNumber) {
    return json(
      {
        error:
          "Nie znaleziono pięciocyfrowego numeru rozpoczynającego się od #."
      },
      400
    );
  }


  /* -------------------------------------------------------
     ODCZYT ŹRÓDŁOWEGO XLSX
     ------------------------------------------------------- */

  const sourceBuffer =
    await source.file.async(
      "arraybuffer"
    );

  const sourceZip =
    await JSZip.loadAsync(
      sourceBuffer
    );


  const sourceWorkbook =
    await readWorkbook(sourceZip);


  /* -------------------------------------------------------
     ODCZYT INFORMACJI O PRZEBIEGACH
     ------------------------------------------------------- */

  const runInfo =
    await extractRunInformation(
      sourceZip,
      sourceWorkbook
    );


  /* -------------------------------------------------------
     POBIERAMY MACIERZYSTY SKOROSZYT
     ------------------------------------------------------- */

  const templateObject =
    await env.STORAGE.get(
      TEMPLATE_KEY
    );

  if (!templateObject) {

    return json(
      {
        error:
          `Nie znaleziono ${TEMPLATE_KEY} w R2.`
      },
      500
    );

  }


  const templateBuffer =
    await templateObject.arrayBuffer();


  /*
   * KLUCZOWA RZECZ:
   *
   * Otwieramy macierzysty XLSX jako ZIP.
   *
   * Nie przebudowujemy skoroszytu biblioteką XLSX.
   * Dzięki temu zachowujemy formatowanie.
   */

  const resultZip =
    await JSZip.loadAsync(
      templateBuffer
    );


  /* -------------------------------------------------------
     ZNAJDUJEMY ARKUSZ ARKUSZ1
     ------------------------------------------------------- */

  const resultWorkbook =
    await readWorkbook(
      resultZip
    );

  const resultSheetPath =
    await findSheetPath(
      resultZip,
      resultWorkbook,
      "Arkusz1"
    );

  if (!resultSheetPath) {

    return json(
      {
        error:
          "W macierzystym skoroszycie nie znaleziono Arkusz1."
      },
      500
    );

  }


  /* -------------------------------------------------------
     XML ARKUSZA
     ------------------------------------------------------- */

  let resultSheetXml =
    await getXml(
      resultZip,
      resultSheetPath
    );


  /* -------------------------------------------------------
     B2 = #21482
     ------------------------------------------------------- */

  resultSheetXml =
    setCellValue(
      resultSheetXml,
      "B2",
      runNumber
    );


  /* -------------------------------------------------------
     MUFY + SPRAWY
     ------------------------------------------------------- */

  resultSheetXml =
    writeMuffs(
      resultSheetXml,
      runInfo.splicedMuffs
    );


  /* -------------------------------------------------------
     ZAPIS ARKUSZA
     ------------------------------------------------------- */

  resultZip.file(
    resultSheetPath,
    resultSheetXml
  );


  /* -------------------------------------------------------
     GENERUJEMY GOTOWY XLSX
     ------------------------------------------------------- */

  const resultBuffer =
    await resultZip.generateAsync({
      type: "arraybuffer",
      compression: "STORE"
    });


  /* -------------------------------------------------------
     NAZWA KATALOGU ZLECENIA
     ------------------------------------------------------- */

  const rootFolder =
    makeRootFolder(
      order.orderNumber,
      order.description
    );


  /* -------------------------------------------------------
     ZAPIS ŹRÓDŁOWEJ PACZKI
     ------------------------------------------------------- */

  const id =
    crypto.randomUUID();

  const inputKey =
    `jobs/${id}/input/${sanitizeFileName(
      uploadedZip.name
    )}`;


  await env.STORAGE.put(
    inputKey,
    zipBuffer,
    {
      httpMetadata: {
        contentType:
          "application/zip"
      }
    }
  );


  /* -------------------------------------------------------
     KATALOGI
     ------------------------------------------------------- */

  for (
    const folder of FOLDERS
  ) {

    await env.STORAGE.put(
      `jobs/${id}/output/${rootFolder}/${folder}/`,
      ""
    );

  }


  /* -------------------------------------------------------
     WYNIK:
     #21482.xlsx
     ------------------------------------------------------- */

  const resultFileName =
    `${runNumber}.xlsx`;

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
     RAPORT ZADANIA
     ------------------------------------------------------- */

  const job = {

    id,

    status:
      "COMPLETED",

    input: {
      rawLine:
        orderLine,

      orderNumber:
        order.orderNumber,

      date:
        order.date,

      operator:
        order.operator,

      code:
        order.code,

      description:
        order.description
    },

    source: {
      file:
        source.name,

      runNumber
    },

    extracted:
      runInfo,

    output: {
      rootFolder,

      workbook:
        resultKey
    },

    createdAt:
      new Date().toISOString()

  };


  await env.STORAGE.put(
    `jobs/${id}/job.json`,
    JSON.stringify(
      job,
      null,
      2
    ),
    {
      httpMetadata: {
        contentType:
          "application/json"
      }
    }
  );


  return json(
    job,
    200
  );
}


/* =========================================================
   PARSOWANIE WIERSZA ZLECENIA
   ========================================================= */

function parseOrderLine(line) {

  const columns =
    line
      .split(/\t+/)
      .map(
        x => x.trim()
      );

  if (
    columns.length < 5
  ) {

    throw new Error(
      "Wiersz zlecenia musi zawierać minimum 5 kolumn."
    );

  }

  return {

    orderNumber:
      columns[0],

    date:
      columns[1],

    operator:
      columns[2],

    code:
      columns[3],

    description:
      columns
        .slice(4)
        .join(" ")
        .trim()

  };
}


/* =========================================================
   SZUKANIE PLIKU XLSX Z #
   ========================================================= */

async function findSourceXlsx(zip) {

  let result =
    null;

  zip.forEach(
    (path, file) => {

      if (result) {
        return;
      }

      if (file.dir) {
        return;
      }

      const name =
        path
          .split("/")
          .pop();

      if (
        name.startsWith("#") &&
        name
          .toLowerCase()
          .endsWith(".xlsx")
      ) {

        result = {
          name,
          file
        };

      }

    }
  );

  return result;
}


/* =========================================================
   # + 5 CYFR
   ========================================================= */

function extractFiveDigitNumber(
  filename
) {

  const match =
    filename.match(
      /#(\d{5})(?!\d)/
    );

  if (!match) {
    return null;
  }

  return `#${match[1]}`;
}


/* =========================================================
   WORKBOOK.XML
   ========================================================= */

async function readWorkbook(zip) {

  const file =
    zip.file(
      "xl/workbook.xml"
    );

  if (!file) {

    throw new Error(
      "Brak xl/workbook.xml."
    );

  }

  return await file.async(
    "string"
  );
}


/* =========================================================
   ZNALEZIENIE ŚCIEŻKI ARKUSZA
   ========================================================= */

async function findSheetPath(
  zip,
  workbookXml,
  sheetName
) {

  const sheetRegex =
    new RegExp(
      `<sheet\\b[^>]*name="${escapeRegex(sheetName)}"[^>]*r:id="([^"]+)"[^>]*/?>`,
      "i"
    );

  const match =
    workbookXml.match(
      sheetRegex
    );

  if (!match) {
    return null;
  }

  const relId =
    match[1];


  const relsFile =
    zip.file(
      "xl/_rels/workbook.xml.rels"
    );

  if (!relsFile) {
    return null;
  }

  const relsXml =
    await relsFile.async(
      "string"
    );


  const relRegex =
    new RegExp(
      `<Relationship\\b[^>]*Id="${escapeRegex(relId)}"[^>]*Target="([^"]+)"[^>]*/?>`,
      "i"
    );

  const relMatch =
    relsXml.match(
      relRegex
    );

  if (!relMatch) {
    return null;
  }


  let target =
    relMatch[1];

  target =
    target.replace(
      /^\/+/,
      ""
    );

  if (
    !target.startsWith("xl/")
  ) {
    target =
      `xl/${target}`;
  }

  return target;
}


/* =========================================================
   XML Z PLIKU
   ========================================================= */

async function getXml(
  zip,
  path
) {

  const file =
    zip.file(path);

  if (!file) {

    throw new Error(
      `Nie znaleziono ${path}.`
    );

  }

  return await file.async(
    "string"
  );
}


/* =========================================================
   USTAWIANIE WARTOŚCI KOMÓRKI
   ========================================================= */

function setCellValue(
  sheetXml,
  cellReference,
  value
) {

  const escaped =
    xmlEscape(value);


  const cellRegex =
    new RegExp(
      `<c\\b([^>]*\\br="${escapeRegex(cellReference)}"[^>]*)>[\\s\\S]*?<\\/c>`,
      "i"
    );


  const existing =
    sheetXml.match(
      cellRegex
    );


  if (existing) {

    const attributes =
      existing[1]
        .replace(
          /\s+t="[^"]*"/gi,
          ""
        );

    const newCell =
      `<c${attributes} t="inlineStr"><is><t>${escaped}</t></is></c>`;

    return sheetXml.replace(
      cellRegex,
      newCell
    );

  }


  /*
   * Jeżeli B2 nie istnieje,
   * dodajemy ją do wiersza 2.
   */

  const rowRegex =
    /<row\b([^>]*\br="2"[^>]*)>/i;

  if (
    rowRegex.test(sheetXml)
  ) {

    const newCell =
      `<c r="${cellReference}" t="inlineStr"><is><t>${escaped}</t></is></c>`;

    return sheetXml.replace(
      rowRegex,
      `$&${newCell}`
    );

  }


  throw new Error(
    `Nie znaleziono wiersza 2 dla komórki ${cellReference}.`
  );
}


/* =========================================================
   MUFY
   ========================================================= */

function writeMuffs(
  sheetXml,
  muffs
) {

  /*
   * F = kolumna 6.
   *
   * XLSX zapisuje kolumny:
   * F, G, H, I...
   */

  for (
    let i = 0;
    i < muffs.length;
    i++
  ) {

    const column =
      numberToColumn(
        6 + i
      );


    const nameCell =
      `${column}2`;

    const spliceCell =
      `${column}4`;


    sheetXml =
      setCellValue(
        sheetXml,
        nameCell,
        muffs[i].name
      );


    sheetXml =
      setNumericCellValue(
        sheetXml,
        spliceCell,
        muffs[i].splices
      );

  }


  return sheetXml;
}


/* =========================================================
   LICZBA W KOMÓRCE
   ========================================================= */

function setNumericCellValue(
  sheetXml,
  cellReference,
  value
) {

  const cellRegex =
    new RegExp(
      `<c\\b([^>]*\\br="${escapeRegex(cellReference)}"[^>]*)>[\\s\\S]*?<\\/c>`,
      "i"
    );


  const existing =
    sheetXml.match(
      cellRegex
    );


  if (existing) {

    const attributes =
      existing[1]
        .replace(
          /\s+t="[^"]*"/gi,
          ""
        );


    const newCell =
      `<c${attributes}><v>${Number(value) || 0}</v></c>`;


    return sheetXml.replace(
      cellRegex,
      newCell
    );

  }


  const rowNumber =
    cellReference.match(
      /\d+$/
    )[0];


  const rowRegex =
    new RegExp(
      `<row\\b([^>]*\\br="${rowNumber}"[^>]*)>`,
      "i"
    );


  if (
    rowRegex.test(sheetXml)
  ) {

    const newCell =
      `<c r="${cellReference}"><v>${Number(value) || 0}</v></c>`;

    return sheetXml.replace(
      rowRegex,
      `$&${newCell}`
    );

  }


  throw new Error(
    `Nie znaleziono wiersza ${rowNumber}.`
  );
}


/* =========================================================
   ODCZYT INFORMACJI O PRZEBIEGACH
   ========================================================= */

async function extractRunInformation(
  zip,
  workbookXml
) {

  const sharedStrings =
    await readSharedStrings(
      zip
    );


  const sheetPaths =
    await findAllSheetPaths(
      zip,
      workbookXml
    );


  const result = {
    splicedMuffs: []
  };


  for (
    const sheetPath of sheetPaths
  ) {

    const xml =
      await getXml(
        zip,
        sheetPath
      );


    const rows =
      parseSheetRows(
        xml,
        sharedStrings
      );


    const text =
      rows
        .map(
          row =>
            row.values.join(" | ")
        )
        .join("\n");


    if (
      !text
        .toLowerCase()
        .includes(
          "informacje o przebiegach"
        )
    ) {

      continue;

    }


    /*
     * Na razie analizujemy tylko arkusz,
     * w którym występuje sekcja.
     */

    for (
      let i = 0;
      i < rows.length;
      i++
    ) {

      const row =
        rows[i];


      const rowText =
        row.values
          .join(" | ");


      const lower =
        rowText.toLowerCase();


      /*
       * "nie istnieje - przecina tube"
       * musi być sprawdzane jako pierwsze.
       */

      if (
        lower.includes(
          "nie istnieje - przecina tube"
        )
      ) {

        const name =
          findMuffNameFromRow(
            row
          );


        if (name) {

          result.splicedMuffs.push({

            name,

            splices:
              12,

            tubeCut:
              true,

            prepareCableEnd:
              1

          });

        }

        continue;
      }


      if (
        lower.includes(
          "nie istnieje"
        )
      ) {

        const name =
          findMuffNameFromRow(
            row
          );


        if (name) {

          const splices =
            findSpliceCount(
              row,
              rows,
              i
            );


          result.splicedMuffs.push({

            name,

            splices,

            tubeCut:
              false,

            prepareCableEnd:
              0

          });

        }

      }

    }


    break;

  }


  return result;
}


/* =========================================================
   WSPÓLNE ARKUSZE
   ========================================================= */

async function findAllSheetPaths(
  zip,
  workbookXml
) {

  const paths = [];


  const relsFile =
    zip.file(
      "xl/_rels/workbook.xml.rels"
    );


  if (!relsFile) {
    return paths;
  }


  const relsXml =
    await relsFile.async(
      "string"
    );


  const sheetRegex =
    /<sheet\b[^>]*r:id="([^"]+)"[^>]*\/?>/gi;


  let match;


  while (
    (match =
      sheetRegex.exec(
        workbookXml
      ))
  ) {

    const relId =
      match[1];


    const relRegex =
      new RegExp(
        `<Relationship\\b[^>]*Id="${escapeRegex(relId)}"[^>]*Target="([^"]+)"[^>]*/?>`,
        "i"
      );


    const rel =
      relsXml.match(
        relRegex
      );


    if (!rel) {
      continue;
    }


    let target =
      rel[1]
        .replace(
          /^\/+/,
          ""
        );


    if (
      !target.startsWith("xl/")
    ) {

      target =
        `xl/${target}`;

    }


    paths.push(target);

  }


  return paths;
}


/* =========================================================
   SHARED STRINGS
   ========================================================= */

async function readSharedStrings(
  zip
) {

  const file =
    zip.file(
      "xl/sharedStrings.xml"
    );


  if (!file) {
    return [];
  }


  const xml =
    await file.async(
      "string"
    );


  const strings = [];


  const regex =
    /<si\b[\s\S]*?<\/si>/gi;


  let match;


  while (
    (match =
      regex.exec(xml))
  ) {

    const si =
      match[0];


    const texts = [];


    const textRegex =
      /<t\b[^>]*>([\s\S]*?)<\/t>/gi;


    let t;


    while (
      (t =
        textRegex.exec(si))
    ) {

      texts.push(
        xmlDecode(
          t[1]
        )
      );

    }


    strings.push(
      texts.join("")
    );

  }


  return strings;
}


/* =========================================================
   WIERSZE ARKUSZA
   ========================================================= */

function parseSheetRows(
  xml,
  sharedStrings
) {

  const rows = [];


  const rowRegex =
    /<row\b[^>]*\br="(\d+)"[^>]*>[\s\S]*?<\/row>/gi;


  let rowMatch;


  while (
    (rowMatch =
      rowRegex.exec(xml))
  ) {

    const rowNumber =
      Number(
        rowMatch[1]
      );


    const rowXml =
      rowMatch[0];


    const values = [];


    const cellRegex =
      /<c\b([^>]*)>([\s\S]*?)<\/c>/gi;


    let cellMatch;


    while (
      (cellMatch =
        cellRegex.exec(
          rowXml
        ))
    ) {

      const attributes =
        cellMatch[1];


      const content =
        cellMatch[2];


      const refMatch =
        attributes.match(
          /\br="([A-Z]+\d+)"/i
        );


      if (!refMatch) {
        continue;
      }


      const ref =
        refMatch[1];


      const typeMatch =
        attributes.match(
          /\bt="([^"]+)"/i
        );


      const type =
        typeMatch
          ? typeMatch[1]
          : "";


      let value = "";


      const valueMatch =
        content.match(
          /<v\b[^>]*>([\s\S]*?)<\/v>/i
        );


      if (valueMatch) {

        value =
          valueMatch[1];

      }


      if (
        type === "s"
      ) {

        const index =
          Number(value);

        value =
          sharedStrings[index] || "";

      }

      else if (
        type === "inlineStr"
      ) {

        const texts = [];


        const textRegex =
          /<t\b[^>]*>([\s\S]*?)<\/t>/gi;


        let t;


        while (
          (t =
            textRegex.exec(
              content
            ))
        ) {

          texts.push(
            xmlDecode(
              t[1]
            )
          );

        }


        value =
          texts.join("");

      }

      else {

        value =
          xmlDecode(
            value
          );

      }


      values.push({
        ref,
        value
      });

    }


    rows.push({
      rowNumber,
      values:
        values.map(
          x => x.value
        ),
      cells:
        values
    });

  }


  return rows;
}


/* =========================================================
   NAZWA MUFY
   ========================================================= */

function findMuffNameFromRow(
  row
) {

  for (
    const cell of row.cells
  ) {

    const value =
      String(
        cell.value || ""
      ).trim();


    if (!value) {
      continue;
    }


    const lower =
      value.toLowerCase();


    if (
      lower.includes(
        "nie istnieje"
      )
    ) {
      continue;
    }


    if (
      lower.includes(
        "liczba spawów"
      )
    ) {
      continue;
    }


    if (
      lower.includes(
        "informacje o przebiegach"
      )
    ) {
      continue;
    }


    return value;

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
   * Najpierw ten sam wiersz.
   */

  for (
    let i = 0;
    i < row.cells.length;
    i++
  ) {

    const value =
      String(
        row.cells[i].value || ""
      );


    if (
      /liczba spawów pomiędzy portem a włóknem/i
        .test(value)
    ) {

      const next =
        row.cells[i + 1];


      if (next) {

        const number =
          Number(
            next.value
          );


        if (
          Number.isFinite(
            number
          )
        ) {

          return number;

        }

      }

    }

  }


  /*
   * Następnie kilka kolejnych wierszy.
   */

  for (
    let offset = 1;
    offset <= 5;
    offset++
  ) {

    const candidate =
      rows[rowIndex + offset];


    if (!candidate) {
      break;
    }


    for (
      let i = 0;
      i < candidate.cells.length;
      i++
    ) {

      const value =
        String(
          candidate.cells[i].value || ""
        );


      if (
        /liczba spawów pomiędzy portem a włóknem/i
          .test(value)
      ) {

        const next =
          candidate.cells[i + 1];


        if (next) {

          const number =
            Number(
              next.value
            );


          if (
            Number.isFinite(
              number
            )
          ) {

            return number;

          }

        }

      }

    }

  }


  /*
   * Nie zgadujemy.
   */

  return 0;
}


/* =========================================================
   NUMER KOLUMNY → LITERA
   ========================================================= */

function numberToColumn(
  number
) {

  let result = "";


  while (
    number > 0
  ) {

    const remainder =
      (number - 1) % 26;


    result =
      String.fromCharCode(
        65 + remainder
      ) +
      result;


    number =
      Math.floor(
        (number - 1) / 26
      );

  }


  return result;
}


/* =========================================================
   NAZWA KATALOGU
   ========================================================= */

function makeRootFolder(
  orderNumber,
  description
) {

  return (
    orderNumber
      .replace(
        /\//g,
        "_"
      ) +
    " " +
    description
  );
}


/* =========================================================
   BEZPIECZNA NAZWA PLIKU
   ========================================================= */

function sanitizeFileName(
  name
) {

  return name.replace(
    /[^a-zA-Z0-9._-]/g,
    "_"
  );
}


/* =========================================================
   XML ESCAPE
   ========================================================= */

function xmlEscape(
  value
) {

  return String(value)
    .replace(
      /&/g,
      "&amp;"
    )
    .replace(
      /</g,
      "&lt;"
    )
    .replace(
      />/g,
      "&gt;"
    )
    .replace(
      /"/g,
      "&quot;"
    )
    .replace(
      /'/g,
      "&apos;"
    );

}


/* =========================================================
   XML DECODE
   ========================================================= */

function xmlDecode(
  value
) {

  return String(value)
    .replace(
      /&lt;/g,
      "<"
    )
    .replace(
      /&gt;/g,
      ">"
    )
    .replace(
      /&quot;/g,
      '"'
    )
    .replace(
      /&apos;/g,
      "'"
    )
    .replace(
      /&amp;/g,
      "&"
    );

}


/* =========================================================
   ESCAPE REGEX
   ========================================================= */

function escapeRegex(
  value
) {

  return String(value)
    .replace(
      /[.*+?^${}()|[\]\\]/g,
      "\\$&"
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
   STRONA
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

<form
  id="jobForm"
>

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

<button
  type="submit"
>

PRZYJMIJ ZLECENIE

</button>

</form>

<div id="status"></div>

<script>

const form =
  document.getElementById(
    "jobForm"
  );

const status =
  document.getElementById(
    "status"
  );

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

        "Numer: " +
        data.source.runNumber +

        "\\n\\nPlik:\\n" +
        data.output.workbook +

        "\\n\\nMufy:\\n" +

        data.extracted
          .splicedMuffs
          .map(
            muff =>
              muff.name +
              " → " +
              muff.splices +
              " spawów"
          )
          .join("\\n");

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
