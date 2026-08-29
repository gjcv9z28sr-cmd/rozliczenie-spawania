export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    if (request.method === "GET" && url.pathname === "/") {
      return new Response(page(), {
        headers: { "content-type": "text/html; charset=utf-8" },
      });
    }

    if (request.method === "POST" && url.pathname === "/api/job") {
      return createJob(request, env);
    }

    if (request.method === "GET" && url.pathname === "/api/job") {
      return getJob(url.searchParams.get("id"), env);
    }

    return new Response("Not found", { status: 404 });
  },
};

async function createJob(request, env) {
  const form = await request.formData();
  const orderNumber = String(form.get("orderNumber") || "").trim();
  const archive = form.get("archive");

  if (!orderNumber) {
    return json({ error: "Brak numeru zlecenia." }, 400);
  }

  if (!(archive instanceof File)) {
    return json({ error: "Nie wybrano paczki ZIP." }, 400);
  }

  if (!archive.name.toLowerCase().endsWith(".zip")) {
    return json({ error: "Plik musi mieć rozszerzenie .zip." }, 400);
  }

  const id = crypto.randomUUID();
  const safeName = archive.name.replace(/[^a-zA-Z0-9._-]/g, "_");
  const inputKey = `jobs/${id}/input/${safeName}`;

  await env.STORAGE.put(inputKey, archive.stream(), {
    httpMetadata: { contentType: "application/zip" },
    customMetadata: {
      orderNumber,
      originalName: archive.name,
    },
  });

  const job = {
    id,
    orderNumber,
    inputKey,
    status: "RECEIVED",
    createdAt: new Date().toISOString(),
  };

  await env.STORAGE.put(
    `jobs/${id}/job.json`,
    JSON.stringify(job, null, 2),
    {
      httpMetadata: {
        contentType: "application/json",
      },
    }
  );

  return json(job, 202);
}

async function getJob(id, env) {
  if (!id) {
    return json({ error: "Brak ID zadania." }, 400);
  }

  const object = await env.STORAGE.get(`jobs/${id}/job.json`);

  if (!object) {
    return json({ error: "Nie znaleziono zadania." }, 404);
  }

  return new Response(await object.text(), {
    headers: {
      "content-type": "application/json; charset=utf-8",
    },
  });
}

function json(value, status = 200) {
  return new Response(JSON.stringify(value, null, 2), {
    status,
    headers: {
      "content-type": "application/json; charset=utf-8",
    },
  });
}

function page() {
  return `<!doctype html>
<html lang="pl">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Agent2 — rozliczenie spawania</title>

<style>
body {
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
  max-width: 650px;
  margin: 0 auto;
  padding: 24px;
}

h1 {
  font-size: 26px;
}

label {
  display: block;
  margin-top: 20px;
  font-weight: 600;
}

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

button {
  margin-top: 24px;
  font-weight: 700;
}

#status {
  margin-top: 24px;
  white-space: pre-wrap;
}
</style>

</head>

<body>

<h1>Agent2 — rozliczenie spawania</h1>

<form id="jobForm">

<label>
Numer zlecenia

<input
name="orderNumber"
placeholder="np. 084/W/2026"
required
>

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
WYŚLIJ PACZKĘ
</button>

</form>

<div id="status"></div>

<script>

const form = document.getElementById("jobForm");
const status = document.getElementById("status");

form.addEventListener("submit", async (event) => {

  event.preventDefault();

  status.textContent = "Wysyłanie paczki...";

  try {

    const response = await fetch("/api/job", {
      method: "POST",
      body: new FormData(form)
    });

    const data = await response.json();

    if (!response.ok) {
      throw new Error(data.error || "Wystąpił błąd.");
    }

    status.textContent =
      "Paczka przyjęta.\\n\\n" +
      "ID zadania: " + data.id + "\\n" +
      "Status: " + data.status;

  } catch (error) {

    status.textContent =
      "Błąd: " + error.message;

  }

});

</script>

</body>
</html>`;
}
