// Issues short-lived, scope-limited tokens so the browser can upload a
// file directly to Vercel Blob storage, without the file ever passing
// through a Vercel Function. Vercel Functions reject any request body
// over 4.5MB regardless of configuration -- this is the documented way
// around that limit. The uploaded file's resulting Blob URL is sent by
// the browser straight to the Flask app afterward; this endpoint never
// sees file contents, only issues the token that authorizes the upload.
const { handleUpload } = require("@vercel/blob/client");

module.exports = async (req, res) => {
  if (req.method !== "POST") {
    res.status(405).json({ error: "Method not allowed" });
    return;
  }
  try {
    const jsonResponse = await handleUpload({
      body: req.body,
      request: req,
      onBeforeGenerateToken: async () => ({
        allowedContentTypes: [
          "application/vnd.openxmlformats-officedocument.presentationml.presentation",
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          "application/vnd.ms-excel",
          "text/csv",
        ],
        addRandomSuffix: true,
        maximumSizeInBytes: 200 * 1024 * 1024,
      }),
      // No onUploadCompleted: the browser reports the resulting Blob URL
      // straight to the Flask route itself once the upload finishes,
      // rather than Vercel Blob calling back to a webhook here.
    });
    res.status(200).json(jsonResponse);
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
};
