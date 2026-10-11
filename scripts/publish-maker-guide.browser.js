export async function publishMakerGuide({
  sitePath,
  pagePath,
  pageItemId,
  expectedTitle,
  expectedOwnerEmail,
  guideHtml,
}) {
  if (!sitePath?.startsWith("/") || !pagePath?.startsWith(`${sitePath}/SitePages/`) ||
      !Number.isInteger(pageItemId) || pageItemId < 1 || !guideHtml?.trim()) {
    throw new Error("A site path, SitePages page path, positive item ID and rendered guide are required.");
  }
  if (!location.hostname.endsWith(".sharepoint.com")) {
    throw new Error("Run this publisher on the signed-in SharePoint site.");
  }
  const headers = { Accept: "application/json;odata=nometadata" };
  const read = async (path) => {
    const response = await fetch(path, { headers });
    if (!response.ok) throw new Error(`SharePoint read failed (${response.status}). Sign in and verify page access.`);
    return { response, data: await response.json() };
  };
  const user = await read(`${sitePath}/_api/web/currentuser?$select=Email`);
  if (user.data.Email?.toLowerCase() !== expectedOwnerEmail?.toLowerCase()) {
    throw new Error("The signed-in SharePoint account is not the deployment owner.");
  }
  const itemUrl = `${sitePath}/_api/web/lists/getbytitle('Site Pages')/items(${pageItemId})`;
  const { response, data } = await read(`${itemUrl}?$select=Id,Title,CanvasContent1`);
  if (data.Title !== expectedTitle) throw new Error("The target page title does not match.");
  const etag = response.headers.get("ETag") || data["@odata.etag"] || data["odata.etag"];
  if (!etag) throw new Error("The page has no concurrency token.");
  const canvas = new DOMParser().parseFromString(data.CanvasContent1, "text/html");
  const controls = [...canvas.querySelectorAll("[data-sp-controldata]")];
  const guideId = "d820b388-21ba-4d57-8e2f-63fbf99b8f45";
  const controlData = (node) => JSON.parse(node.getAttribute("data-sp-controldata"));
  const list = controls.find((node) => controlData(node).controlType === 3);
  if (!list) throw new Error("The expected list web part is missing; no page changes were made.");
  let guide = controls.find((node) => controlData(node).id === guideId);
  if (!guide) {
    guide = canvas.createElement("div");
    guide.setAttribute("data-sp-canvascontrol", "");
    guide.setAttribute("data-sp-canvasdataversion", "1.0");
    const position = {
      ...controlData(list).position,
      controlIndex: Math.max(...controls.map((node) => controlData(node).position?.controlIndex || 0)) + 1,
    };
    guide.setAttribute("data-sp-controldata", JSON.stringify({ controlType: 4, id: guideId, position }));
    list.parentElement.appendChild(guide);
  }
  const content = canvas.createElement("div");
  content.setAttribute("data-sp-rte", "");
  content.innerHTML = guideHtml;
  guide.replaceChildren(content);
  const digestResponse = await fetch(`${sitePath}/_api/contextinfo`, { method: "POST", headers });
  if (!digestResponse.ok) throw new Error("SharePoint request digest could not be acquired.");
  const digestData = await digestResponse.json();
  const digest = digestData.FormDigestValue || digestData.d?.GetContextWebInformation?.FormDigestValue;
  if (!digest) throw new Error("SharePoint returned no request digest.");
  const save = await fetch(itemUrl, {
    method: "POST",
    headers: {
      ...headers,
      "Content-Type": "application/json;odata=nometadata",
      "X-RequestDigest": digest,
      "IF-MATCH": etag,
      "X-HTTP-Method": "MERGE",
    },
    body: JSON.stringify({ CanvasContent1: canvas.body.innerHTML }),
  });
  if (!save.ok) throw new Error(`Page save failed (${save.status}); reload before retrying.`);
  const escapedPath = pagePath.replaceAll("'", "''");
  const publish = await fetch(
    `${sitePath}/_api/web/GetFileByServerRelativeUrl('${escapedPath}')/Publish(comment='Updated maker guide')`,
    { method: "POST", headers: { ...headers, "X-RequestDigest": digest } },
  );
  if (!publish.ok) throw new Error(`The guide was saved but page publication failed (${publish.status}).`);
  const verified = await read(`${itemUrl}?$select=CanvasContent1`);
  const verifiedCanvas = new DOMParser().parseFromString(verified.data.CanvasContent1, "text/html");
  const savedControls = [...verifiedCanvas.querySelectorAll("[data-sp-controldata]")];
  if (!savedControls.some((node) => controlData(node).id === guideId) ||
      controls.some((node) => !savedControls.some((saved) => controlData(saved).id === controlData(node).id))) {
    throw new Error("Published page verification failed; inspect the guide and existing controls.");
  }
  return { published: true, pageItemId, controlsPreserved: controls.length, guideCharacters: guideHtml.length };
}
