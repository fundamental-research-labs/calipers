/* Calipers runner: fetch corpus Office.js from /job, Excel.run it, POST /done. */
fetch("/loaded").catch(() => {});
Office.onReady(async (info) => {
  const status = document.getElementById("status");
  const set = (t) => {
    if (status) status.textContent = t;
  };
  if (!info || info.host !== Office.HostType.Excel) {
    await postDone({ ok: false, error: "not excel" });
    set("not excel");
    return;
  }
  let job;
  try {
    const res = await fetch("/job");
    job = await res.json();
  } catch (e) {
    await postDone({ ok: false, error: "fetch /job: " + String(e) });
    set("job fetch failed");
    return;
  }
  try {
    await runCorpusScript(job.script);
    await postDone({ ok: true });
    set("done");
  } catch (e) {
    const msg = e && e.message ? e.message : String(e);
    await postDone({ ok: false, error: msg });
    set("error: " + msg);
  }
});

async function runCorpusScript(script) {
  if (typeof script !== "string" || !script.trim()) {
    throw new Error("empty script");
  }
  // Corpus files use top-level await Excel.run(...).
  const AsyncFunction = Object.getPrototypeOf(async function () {}).constructor;
  const fn = new AsyncFunction(script);
  await fn();
}

async function postDone(body) {
  await fetch("/done", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
}
