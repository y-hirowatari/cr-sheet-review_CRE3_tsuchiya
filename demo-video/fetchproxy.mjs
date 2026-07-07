// Shared helper: route all page requests through node fetch (works via agent proxy)
export async function installFetchRoute(context) {
  await context.route('**/*', async (route) => {
    const req = route.request();
    try {
      const headers = { ...req.headers() };
      delete headers['accept-encoding'];
      const resp = await fetch(req.url(), {
        method: req.method(),
        headers,
        body: ['GET','HEAD'].includes(req.method()) ? undefined : req.postDataBuffer(),
        redirect: 'manual',
      });
      const respHeaders = {};
      resp.headers.forEach((v, k) => {
        if (!['content-encoding','content-length','transfer-encoding','connection'].includes(k)) respHeaders[k] = v;
      });
      const body = Buffer.from(await resp.arrayBuffer());
      await route.fulfill({ status: resp.status, headers: respHeaders, body });
    } catch (e) {
      await route.abort('failed').catch(() => {});
    }
  });
}
