export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    if (request.method === "GET" && url.pathname === "/")
      return new Response("Agent2 OK");
    if (request.method === "POST" && url.pathname === "/api/job")
      return createJob(request, env);
    return new Response("Not found", { status: 404 });
  }
};

async function createJob(request, env) {
  const form = await request.formData();
  const orderNumber = String(form.get("orderNumber") || "").trim();
  const archive = form.get("archive");

  if (!orderNumber) return json({error:"Brak numeru zlecenia."},400);
  if (!(archive instanceof File)) return json({error:"Brak paczki ZIP."},400);
  if (!archive.name.toLowerCase().endsWith(".zip"))
    return json({error:"Paczka musi być ZIP."},400);

  const id = crypto.randomUUID();
  const name = archive.name.replace(/[^a-zA-Z0-9._-]/g, "_");
  const key = `jobs/${id}/input/${name}`;

  await env.STORAGE.put(key, archive.stream(), {
    httpMetadata:{contentType:"application/zip"},
    customMetadata:{orderNumber, originalName:archive.name}
  });

  await env.STORAGE.put(`jobs/${id}/job.json`,
    JSON.stringify({id,orderNumber,inputKey:key,status:"RECEIVED",createdAt:new Date().toISOString()},null,2),
    {httpMetadata:{contentType:"application/json"}}
  );

  return json({ok:true,id,status:"RECEIVED"},202);
}

function json(value,status=200){
  return new Response(JSON.stringify(value),{
    status,
    headers:{"content-type":"application/json; charset=utf-8"}
  });
}
