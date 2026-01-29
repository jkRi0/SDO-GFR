// jsPDF PDF generation with manual header, footer, and paginated body
// This script is self-contained. Load jsPDF in your HTML, then call window.generatePaginatedPdf().
// No manual copy-paste needed. Robust error handling and comments included.

(function () {
  // Expose as window.generatePaginatedPdf
  window.generatePaginatedPdf = async function generatePaginatedPdf(options = {}) {
    const jsPDF = window.jspdf?.jsPDF;
    if (!jsPDF) {
      alert('jsPDF is not loaded. Please include jsPDF before this script.');
      return;
    }

    // Paths to header/footer images (relative to HTML file)
    const headerUrl = '../assets/header.jpg';
    const footerUrl = '../assets/footer.jpg';

    let headerDataUrl, footerDataUrl;
    try {
      [headerDataUrl, footerDataUrl] = await Promise.all([
        loadImageAsDataURL(headerUrl),
        loadImageAsDataURL(footerUrl)
      ]);
    } catch (e) {
      alert('Failed to load header/footer images. Check paths and CORS.');
      return;
    }

    const doc = new jsPDF({ unit: 'mm', format: 'a4', orientation: 'portrait' });
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();

    // Compute header image aspect ratio and size (wait for image load to avoid race)
    const headerSize = await getImageSize(headerDataUrl);
    const headerImgWidth = headerSize.width || 1;
    const headerImgHeight = headerSize.height || 1;
    const headerAspect = headerImgWidth / headerImgHeight;
    const headerW = Math.min(55, pageWidth - 20);
    const headerH = headerW / headerAspect;
    const headerY = 10;

    // Compute footer image aspect ratio and size (independent, also wait for load)
    const footerSize = await getImageSize(footerDataUrl);
    const footerImgWidth = footerSize.width || 1;
    const footerImgHeight = footerSize.height || 1;
    const footerAspect = footerImgWidth / footerImgHeight;
    const footerW = pageWidth - 20; // full width minus margins
    const footerH = footerW / footerAspect;
    const footerY = pageHeight - footerH - 10;
    const topMargin = 8;
    const bottomMargin = 8;
    const bodyTopY = headerY + headerH + topMargin + 4; // extra 4mm gap
    const bodyBottomY = pageHeight - footerH - bottomMargin;
    const bodyHeight = bodyBottomY - bodyTopY;

    function drawHeaderFooter() {
      // Header image (preserve aspect ratio)
      const headerX = (pageWidth - headerW) / 2;
      doc.addImage(headerDataUrl, 'JPEG', headerX, headerY, headerW, headerH);
      // Footer image (aspect-ratio preserved)
      const footerX = (pageWidth - footerW) / 2;
      doc.addImage(footerDataUrl, 'JPEG', footerX, footerY, footerW, footerH);
      // Footer text
      doc.setFontSize(8);
      doc.setTextColor(107, 114, 128);
      doc.text('Sample output only - not based on real records', 10, footerY - 2);
    }

    drawHeaderFooter();
    let y = bodyTopY;
    doc.setFontSize(12);
    doc.setTextColor(34, 34, 34);

    // Add invisible placeholder body to reserve space (prevents footer stretch)
    doc.setTextColor(255, 255, 255); // invisible text
    const placeholderLines = Math.floor(bodyHeight / 6);
    for (let i = 0; i < placeholderLines; i++) {
      doc.text('.', 12, y);
      y += 6;
    }
    y = bodyTopY; // reset y for overlay

    // Render body sections as images using html2canvas (preserves design)
    const bodySections = document.querySelectorAll('#pdf-content .section');
    if (!bodySections.length) {
      alert('No body content found. Ensure #pdf-content .section exists.');
      return;
    }
    for (const section of bodySections) {
      const canvas = await html2canvas(section, { backgroundColor: '#ffffff', scale: 2 });
      const imgData = canvas.toDataURL('image/png');
      const imgW = pageWidth - 20;
      const imgH = (canvas.height / canvas.width) * imgW;
      // Check if we need a new page for this section
      if (y + imgH > bodyBottomY) {
        doc.addPage();
        drawHeaderFooter();
        // Add placeholder on new page too
        doc.setTextColor(255, 255, 255);
        for (let i = 0; i < placeholderLines; i++) {
          doc.text('.', 12, bodyTopY + i * 6);
        }
        y = bodyTopY;
      }
      doc.addImage(imgData, 'PNG', 10, y, imgW, imgH);
      y += imgH + 6; // extra space between sections
    }

    // Generate PDF blob once
    const pdfBlob = doc.output('blob');

    if (options.preview) {
      // Open PDF in a new tab
      const url = URL.createObjectURL(pdfBlob);
      window.open(url, '_blank');
    } else {
      // Download the PDF
      const url = URL.createObjectURL(pdfBlob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'paginated-report.pdf';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }
  };

  // Helper to load image as Data URL
  function loadImageAsDataURL(url) {
    return new Promise(function (resolve, reject) {
      var img = new window.Image();
      img.crossOrigin = 'anonymous';
      img.onload = function () {
        try {
          var canvas = document.createElement('canvas');
          canvas.width = img.naturalWidth;
          canvas.height = img.naturalHeight;
          var ctx = canvas.getContext('2d');
          ctx.drawImage(img, 0, 0);
          var dataURL = canvas.toDataURL('image/jpeg', 0.98);
          resolve(dataURL);
        } catch (err) {
          reject(err);
        }
      };
      img.onerror = function (e) { reject(e); };
      img.src = url;
    });
  }

  // Helper to get natural image size from a data URL
  function getImageSize(dataUrl) {
    return new Promise(function (resolve, reject) {
      var img = new window.Image();
      img.onload = function () {
        resolve({ width: img.naturalWidth, height: img.naturalHeight });
      };
      img.onerror = function (e) { reject(e); };
      img.src = dataUrl;
    });
  }
})();

// Now you can call window.generatePaginatedPdf() from anywhere after DOM is loaded.
