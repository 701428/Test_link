export async function downloadPdf(elementId, filename) {
  const html2pdf = (await import("html2pdf.js")).default;
  const element = document.getElementById(elementId);
  await html2pdf()
    .set({
      margin: [12, 14, 12, 14],
      filename,
      image: { type: "jpeg", quality: 0.98 },
      html2canvas: { scale: 2, useCORS: true, logging: false },
      jsPDF: { unit: "mm", format: "a4", orientation: "portrait" },
      pagebreak: { mode: ["avoid-all", "css"] },
    })
    .from(element)
    .save();
}
