const fs = require('fs');

try {
  // Read Premier - Marca-02.jpg as Base64 for the PDF logo
  const pdfImageBytes = fs.readFileSync('Premier - Marca-02.jpg');
  const pdfBase64Image = pdfImageBytes.toString('base64');
  const pdfLogoData = 'data:image/jpeg;base64,' + pdfBase64Image;

  // Read current HTML
  let html = fs.readFileSync('index.html', 'utf8');

  // 1. Insert LOGO_PDF_BASE64 right below LOGO_BASE64
  const logoBase64Line = html.match(/const LOGO_BASE64 = "[^"]+";/);
  if (logoBase64Line) {
    const replacement = logoBase64Line[0] + `\n      const LOGO_PDF_BASE64 = "${pdfLogoData}";`;
    html = html.replace(logoBase64Line[0], replacement);
  }

  // 2. Change PDF logo rendering logic to use LOGO_PDF_BASE64 as JPEG
  html = html.replace(
    /if \(LOGO_BASE64\) \{[\s\S]*?doc\.addImage\(LOGO_BASE64, 'PNG', marginX, y, 22, 15\.8\);[\s\S]*?\}/,
    `if (LOGO_PDF_BASE64) {
            try {
              // width: 22mm, height: 22 / 1.3885 = 15.8mm
              doc.addImage(LOGO_PDF_BASE64, 'JPEG', marginX, y, 22, 15.8);
            } catch (err) {
              console.error("Error drawing logo in PDF:", err);
            }
          }`
  );

  // 3. Change highlight line to gold in PDF (RGB 212, 175, 55)
  html = html.replace(
    /doc\.setDrawColor\(232, 118, 26\); \/\/ #E8761A/,
    'doc.setDrawColor(212, 175, 55); // #D4AF37 (Gold)'
  );

  // 4. Center the title in the PDF
  html = html.replace(
    /doc\.text\("FICHA DE CADASTRO DE ADMISSÃO", marginX \+ 26, y \+ 5\);/,
    'doc.text("FICHA DE CADASTRO DE ADMISSÃO", marginX + printableWidth / 2, y + 5, { align: "center" });'
  );

  // 5. Center the title in the HTML header
  const oldHtmlHeader = `<div className="flex flex-col md:flex-row items-center gap-6 text-center md:text-left">
                {LOGO_BASE64 && (
                  <div className="bg-white px-4 py-2 shrink-0 rounded-xl overflow-hidden shadow-lg h-16 md:h-20 flex items-center justify-center">
                    <img src={LOGO_BASE64} alt="Premier Logistics Logo" className="h-full w-auto object-contain" />
                  </div>
                )}
                <div>
                  <span className="text-premier-gold uppercase font-bold text-xs tracking-widest block mb-1">
                    Premier Logistics Gestão Empresarial Ltda
                  </span>
                  <h1 className="text-2xl md:text-3xl font-extrabold tracking-tight">
                    Ficha de Cadastro de Admissão
                  </h1>
                </div>
              </div>`;

  const newHtmlHeader = `<div className="flex flex-col md:flex-row items-center gap-6 w-full">
                {LOGO_BASE64 && (
                  <div className="bg-white px-4 py-2 shrink-0 rounded-xl overflow-hidden shadow-lg h-16 md:h-20 flex items-center justify-center">
                    <img src={LOGO_BASE64} alt="Premier Logistics Logo" className="h-full w-auto object-contain" />
                  </div>
                )}
                <div className="text-center flex-1 md:pr-24">
                  <span className="text-premier-gold uppercase font-bold text-xs tracking-widest block mb-1">
                    Premier Logistics Gestão Empresarial Ltda
                  </span>
                  <h1 className="text-2xl md:text-3xl font-extrabold tracking-tight">
                    Ficha de Cadastro de Admissão
                  </h1>
                </div>
              </div>`;

  if (html.includes(oldHtmlHeader)) {
    html = html.replace(oldHtmlHeader, newHtmlHeader);
  } else {
    // Loose regex replacement for the HTML header block if spacing differs
    html = html.replace(
      /<div className="flex flex-col md:flex-row items-center gap-6 text-center md:text-left">([\s\S]*?)<\/div>([\s\S]*?)<\/div>/,
      `<div className="flex flex-col md:flex-row items-center gap-6 w-full">$1</div><div className="text-center flex-1 md:pr-24">$2</div></div>`
    );
  }

  fs.writeFileSync('index.html', html);
  console.log('Successfully completed logo updates, centered titles and applied gold highlight lines.');

} catch (e) {
  console.error('Error:', e);
}
