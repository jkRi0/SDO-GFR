// jsPDF PDF generation with manual header, footer, and paginated body
// This script is self-contained. Load jsPDF in your HTML, then call window.generatePaginatedPdf().
// No manual copy-paste needed. Robust error handling and comments included.

(function () {
  // Helper to render the full report for one office into an existing doc
  function renderOfficeReport(doc, opts, shared) {
    const {
      pageWidth,
      pageHeight,
      headerDataUrl,
      footerDataUrl,
      headerW,
      headerH,
      headerY,
      footerW,
      footerH,
      footerY,
      bodyTopY,
      bodyBottomY,
      bodyHeight,
    } = shared;

    const { office = '(not specified)', period = '(not specified)', totals = {} } = opts;

    function drawHeaderFooter() {
      // Header image (preserve aspect ratio)
      const headerX = (pageWidth - headerW) / 2;
      doc.addImage(headerDataUrl, 'JPEG', headerX, headerY, headerW, headerH);
      // Footer image (aspect-ratio preserved)
      const footerX = (pageWidth - footerW) / 2;
      doc.addImage(footerDataUrl, 'JPEG', footerX, footerY, footerW, footerH);
      // Divider lines: one below header, one above footer
      const lineMargin = 10; // left/right margin for the lines
      const headerLineY = headerY + headerH + 1.5;
      const footerLineY = footerY - 1.5;
      doc.setDrawColor(0);
      doc.setLineWidth(0.3);
      doc.line(lineMargin, headerLineY, pageWidth - lineMargin, headerLineY);
      doc.line(lineMargin, footerLineY, pageWidth - lineMargin, footerLineY);

      // Footer text (optional, currently disabled)
      doc.setFontSize(8);
      doc.setTextColor(107, 114, 128);
      // doc.text('Sample output only - not based on real records', 10, footerY - 2);
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
    // Move the main title slightly closer to the header by starting
    // a few millimeters above the bodyTopY position.
    y = bodyTopY - 6;

    // BODY CONTENT: title + office/period line + summary table
    doc.setTextColor(34, 34, 34);

    // Centered title
    const title = 'CUSTOMER FEEDBACK REPORT';
    doc.setFontSize(11);
    doc.setFont(undefined, 'bold');
    const titleWidth = doc.getTextWidth(title);
    const titleX = (pageWidth - titleWidth) / 2;
    doc.text(title, titleX, y);
    y += 7; // slightly closer to office line

    // Office and Rating Period line (left and right within a centered block)
    doc.setFontSize(11);
    doc.setFont(undefined, 'normal');

    const officeLabel = 'Office Service Availed: ';
    const officeValue = office;
    const periodLabel = 'Rating Period: ';
    const periodValue = period;

    // Content block width (office line + table) for centering the table.
    // We keep the table at ~120mm wide, centered, but allow the office/period
    // line to use almost the full page width.
    const contentWidth = 120; // table area width in mm
    const contentLeft = (pageWidth - contentWidth) / 2;

    // These will also be used for the table box so everything aligns
    const tableStartX = contentLeft;

    // Left side: office text with underline under the office value.
    // Use a wider left margin than the table so there's more room.
    const officeLineLeftMargin = 15;
    const leftX = officeLineLeftMargin;

    const officeText = officeLabel + officeValue;
    const officeLabelWidth = doc.getTextWidth(officeLabel);
    const officeValueWidth = doc.getTextWidth(officeValue);
    const officeTextWidth = officeLabelWidth + officeValueWidth;

    const periodText = periodLabel + periodValue;
    const periodTextWidth = doc.getTextWidth(periodText);
    const periodLabelWidth = doc.getTextWidth(periodLabel);
    const periodValueWidth = doc.getTextWidth(periodValue);
    const officeLineRightMargin = 15;

    // If office + period would overlap, move period to the next line.
    const totalNeededWidth = officeLineLeftMargin + officeTextWidth + periodTextWidth + officeLineRightMargin;
    if (totalNeededWidth <= pageWidth) {
      // Single-line layout: office on left, rating period on right
      doc.text(officeText, leftX, y);
      const officeValueStartX = leftX + officeLabelWidth;
      doc.line(officeValueStartX, y + 1.5, officeValueStartX + officeValueWidth, y + 1.5);

      const rightX = pageWidth - officeLineRightMargin - periodTextWidth;
      doc.text(periodText, rightX, y);

      // Underline the rating period value (e.g., Jan-Dec)
      const periodValueStartX = rightX + periodLabelWidth;
      doc.line(periodValueStartX, y + 1.5, periodValueStartX + periodValueWidth, y + 1.5);

      y += 5;
    } else {
      // Two-line layout: office on first line, rating period directly under it (left-aligned)
      doc.text(officeText, leftX, y);
      const officeValueStartX = leftX + officeLabelWidth;
      doc.line(officeValueStartX, y + 1.5, officeValueStartX + officeValueWidth, y + 1.5);

      // Move down for the period line and print it under the office text
      y += 5;
      doc.text(periodText, leftX, y);

      // Underline the rating period value when wrapped
      const periodValueStartXWrapped = leftX + periodLabelWidth;
      doc.line(periodValueStartXWrapped, y + 1.5, periodValueStartXWrapped + periodValueWidth, y + 1.5);

      // Extra space before the table when wrapped
      y += 4;
    }

    // Summary table
    doc.setFontSize(10);

    const v = (key) => (typeof totals[key] === 'number' ? totals[key] : 0);

    const sexMale = v('sexMale');
    const sexFemale = v('sexFemale');
    const sexTotal = v('sexTotal');

    const age19Lower = v('age19Lower');
    const age20_34 = v('age20_34');
    const age35_49 = v('age35_49');
    const age50_64 = v('age50_64');
    const ageTotal = v('ageTotal');

    const custBusiness = v('custBusiness');
    const custCitizen = v('custCitizen');
    const custGovernment = v('custGovernment');
    const custTotal = v('custTotal');

    // Use jsPDF AutoTable for proper table rendering
    const head = [[
      'Sex',
      'No.',
      'Age',
      'No.',
      'Customer Type',
      'No.',
    ]];

    const body = [
      ['Male', sexMale, '19-lower:', age19Lower, 'Business', custBusiness],
      ['Female', sexFemale, '20-34:', age20_34, 'Citizen', custCitizen],
      ['', '', '35-49:', age35_49, 'Government', custGovernment],
      ['', '', '50-64:', age50_64, '', ''],
      ['Total', sexTotal, 'Total', ageTotal, 'Total', custTotal],
    ];

    doc.autoTable({
      startY: y,
      head,
      body,
      theme: 'grid', // plain grid: outer + inner borders
      styles: {
        fontSize: 8,
        halign: 'left',
        valign: 'middle',
        cellPadding: 0.8,
        fillColor: [255, 255, 255], // white cells
        textColor: 0,
        lineColor: 0,
        lineWidth: 0.1,
      },
      headStyles: {
        fontStyle: 'bold',
        halign: 'center',
        fillColor: [255, 255, 255], // no blue header
        textColor: 0,
      },
      alternateRowStyles: {
        fillColor: [255, 255, 255], // disable zebra striping
      },
      tableLineWidth: 0.2,
      tableLineColor: 0,
      margin: { left: tableStartX, right: pageWidth - tableStartX - 120 },
      columnStyles: {
        1: { halign: 'center' }, // No. (Sex)
        3: { halign: 'center' }, // No. (Age)
        5: { halign: 'center' }, // No. (Customer Type)
      },
    });

    // SECOND TABLE: Service Availed + Citizen's Charter in a single table
    const serviceRows = Array.isArray(totals.serviceRows) ? totals.serviceRows : [];
    const ccTotals = totals.ccTotals || null;

    const firstTableBottomY = (doc.lastAutoTable && doc.lastAutoTable.finalY) || (y + 20);
    const startY2 = firstTableBottomY + 6;

    const head2 = [[
      'Service availed',
      'No.',
      "Citizen's Charter",
      'Yes',
      'No',
      'Did Not Specify',
    ]];

    const body2 = [];

    // Helper to get CC row data by index 0->CC1, 1->CC2, 2->CC3
    function getCcRowForIndex(idx) {
      if (!ccTotals) return { label: '', yes: '', no: '', dns: '' };
      if (idx === 0 && ccTotals.cc1) {
        return {
          label: 'CC1',
          yes: ccTotals.cc1.yes,
          no: ccTotals.cc1.no,
          dns: ccTotals.cc1.didNotSpecify,
        };
      }
      if (idx === 1 && ccTotals.cc2) {
        return {
          label: 'CC2',
          yes: ccTotals.cc2.yes,
          no: ccTotals.cc2.no,
          dns: ccTotals.cc2.didNotSpecify,
        };
      }
      if (idx === 2 && ccTotals.cc3) {
        return {
          label: 'CC3',
          yes: ccTotals.cc3.yes,
          no: ccTotals.cc3.no,
          dns: ccTotals.cc3.didNotSpecify,
        };
      }
      return { label: '', yes: '', no: '', dns: '' };
    }

    // Set margin for 2nd table
    const secondTableWidth = 140;
    const secondTableMargin = (pageWidth - secondTableWidth) / 2;

    // 2nd table row logic: always 3 rows if <3 services, else all services (first 3 paired with CCs)
    if (serviceRows.length < 3) {
      for (let i = 0; i < 3; i++) {
        const srv = serviceRows[i] || { name: '', count: '' };
        const cc = getCcRowForIndex(i);
        body2.push([
          srv.name,
          srv.count,
          cc.label,
          cc.yes,
          cc.no,
          cc.dns,
        ]);
      }
    } else {
      for (let i = 0; i < serviceRows.length; i++) {
        const srv = serviceRows[i];
        let cc = { label: '', yes: '', no: '', dns: '' };
        if (i < 3) cc = getCcRowForIndex(i);
        body2.push([
          srv.name,
          srv.count,
          cc.label,
          cc.yes,
          cc.no,
          cc.dns,
        ]);
      }
    }
    // Always add total row
    const totalServicesCount = serviceRows.reduce((sum, s) => sum + (s.count || 0), 0);
    body2.push([
      'Total',
      totalServicesCount,
      '', '', '', '',
    ]);

    doc.autoTable({
      startY: startY2,
      head: head2,
      body: body2,
      theme: 'grid',
      styles: {
        fontSize: 8,
        halign: 'left',
        valign: 'middle',
        cellPadding: 0.8,
        fillColor: [255, 255, 255],
        textColor: 0,
        lineColor: 0,
        lineWidth: 0.1,
      },
      headStyles: {
        fontStyle: 'bold',
        halign: 'center',
        fillColor: [255, 255, 255],
        textColor: 0,
      },
      alternateRowStyles: {
        fillColor: [255, 255, 255],
      },
      tableLineWidth: 0.2,
      tableLineColor: 0,
      margin: { left: secondTableMargin, right: secondTableMargin },
      columnStyles: {
        1: { halign: 'center' }, // No.
        3: { halign: 'center' }, // Yes
        4: { halign: 'center' }, // No
        5: { halign: 'center' }, // Did Not Specify
      },
    });

    // THIRD TABLE: CLIENT SATISFACTION (wider, more columns)
    const csRows = Array.isArray(totals.clientSatisfactionRows)
      ? totals.clientSatisfactionRows
      : [];

    const secondTableBottomY = (doc.lastAutoTable && doc.lastAutoTable.finalY) || (startY2 + 20);
    const startY3 = secondTableBottomY + 8;

    // Title above the table
    const csTitle = 'CLIENT SATISFACTION';
    doc.setFontSize(11);
    doc.setFont(undefined, 'bold');
    const csTitleWidth = doc.getTextWidth(csTitle);
    const csTitleX = (pageWidth - csTitleWidth) / 2;
    doc.text(csTitle, csTitleX, startY3);

    const tableStartY3 = startY3 + 4;

    const head3 = [[
      'Survey',
      'Strongly Agree (5)',
      'Agree (4)',
      'Neither Agree nor DisAgree (3)',
      'Disagree (2)',
      'Strongly Disagree (1)',
      'Not Applicable',
      'Total No. of Respondents',
      'Total Rated Score',
      'Ave. Rated Score',
    ]];

    let body3;

    if (csRows.length === 0) {
      // Fallback row when there are no client satisfaction entries
      body3 = [[
        'No client satisfaction data',
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        '',
      ]];
    } else {
      body3 = csRows.map((row) => [
        row.label,
        row.sa5,
        row.a4,
        row.n3,
        row.d2,
        row.sd1,
        row.na,
        row.totalRespondents,
        row.totalRatedScore,
        row.averageScore,
      ]);

      // Compute overall average rating across all SQDs
      let overallTotalScore = 0;
      let overallTotalRespondents = 0;
      csRows.forEach((row) => {
        const tr = typeof row.totalRatedScore === 'number' ? row.totalRatedScore : 0;
        const resp = typeof row.totalRespondents === 'number' ? row.totalRespondents : 0;
        overallTotalScore += tr;
        overallTotalRespondents += resp;
      });
      const overallAverage = overallTotalRespondents > 0
        ? (overallTotalScore / overallTotalRespondents).toFixed(2)
        : '';

      // Add "Overall Ratings" row at the bottom, with overall average in last column
      body3.push([
        'Overall Ratings',
        '',
        '',
        '',
        '',
        '',
        '',
        '',
        '',
        overallAverage,
      ]);
    }

    // Make this table a bit wider than the second one and center it.
    const thirdTableWidth = 180; // close to full width
    const thirdTableMargin = (pageWidth - thirdTableWidth) / 2;

    doc.autoTable({
      startY: tableStartY3,
      head: head3,
      body: body3,
      theme: 'grid',
      styles: {
        fontSize: 7,
        halign: 'left',
        valign: 'middle',
        cellPadding: 0.8,
        fillColor: [255, 255, 255],
        textColor: 0,
        lineColor: 0,
        lineWidth: 0.1,
      },
      headStyles: {
        fontStyle: 'bold',
        halign: 'center',
        fillColor: [255, 255, 255],
        textColor: 0,
      },
      alternateRowStyles: {
        fillColor: [255, 255, 255],
      },
      tableLineWidth: 0.2,
      tableLineColor: 0,
      // Slightly widen the first column (Survey) so its right border moves right
      columnStyles: {
        // Survey column: wider, left-aligned text
        0: { cellWidth: 35, halign: 'left' },
        // All numeric columns: center the numbers
        1: { halign: 'center' },
        2: { halign: 'center' },
        3: { halign: 'center' },
        4: { halign: 'center' },
        5: { halign: 'center' },
        6: { halign: 'center' },
        7: { halign: 'center' },
        8: { halign: 'center' },
        9: { halign: 'center' },
      },
      margin: { left: thirdTableMargin, right: thirdTableMargin },
    });

    // Signature block positioned at the top of the footer area
    // Use a fixed Y based on the footer image position so it always
    // sits just above the footer, regardless of table height.
    let sigY = footerY - 43; // a bit above the footer image

    doc.setFontSize(10);
    doc.setFont(undefined, 'normal');
    // Move the signature block a little towards the center, but still
    // visually anchored under the third table width
    const sigLeft = thirdTableMargin + 20;

    doc.text('Prepared by:', sigLeft, sigY);
    sigY += 8;

    doc.setFont(undefined, 'bold');
    doc.text('CHEM JAYDER M. CABUNGCAL', sigLeft, sigY);
    sigY += 5;

    doc.setFont(undefined, 'normal');
    doc.text('Information Technology Officer I', sigLeft, sigY);
    sigY += 10;

    doc.text('Noted by:', sigLeft, sigY);
    sigY += 8;

    doc.setFont(undefined, 'bold');
    doc.text('CHRISTOPHER R. DIAZ, CESO V', sigLeft, sigY);
    sigY += 5;

    doc.setFont(undefined, 'normal');
    doc.text('Schools Division Superintendent', sigLeft, sigY);
  }

  // Expose as window.generatePaginatedPdf
  window.generatePaginatedPdf = async function generatePaginatedPdf(options = {}) {
    const jsPDF = window.jspdf?.jsPDF;
    if (!jsPDF) {
      alert('jsPDF is not loaded. Please include jsPDF before this script.');
      return;
    }

    // Paths to header/footer images (relative to HTML file)
    const headerUrl = './assets/header.jpg';
    const footerUrl = './assets/footer.jpg';

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

    const shared = {
      pageWidth,
      pageHeight,
      headerDataUrl,
      footerDataUrl,
      headerW,
      headerH,
      headerY,
      footerW,
      footerH,
      footerY,
      bodyTopY,
      bodyBottomY,
      bodyHeight,
    };

    const multiOffices = Array.isArray(options.multiOffices) ? options.multiOffices : null;

    if (multiOffices && multiOffices.length) {
      // Render first office on the first page, then add a page per office
      multiOffices.forEach((entry, idx) => {
        if (idx > 0) {
          doc.addPage();
        }
        renderOfficeReport(doc, entry, shared);
      });
    } else {
      // Single-office mode (existing behavior)
      renderOfficeReport(doc, {
        office: options.office || '(not specified)',
        period: options.period || '(not specified)',
        totals: options.totals || {},
      }, shared);
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

  function getServicesForOffice(office) {
    if (!office) return [];
    const key = String(office).toLowerCase();

    if (key.includes('sds')) {
      return [
        'Feedback/Complaint',
        'Travel authority',
        'Other requests/inquiries',
      ];
    }

    if (key.includes('asds')) {
      return [
        'Other requests/inquiries',
      ];
    }

    if (key.includes('cash') || key.includes('general services') || key.includes('procurement')) {
      return [
        'General Services-related',
        'Cash Advance',
        'Procurement-related',
        'Other requests/inquiries',
      ];
    }

    if (key.includes('personnel')) {
      return [
        'Appointment (new, promotion, transfer, etc.)',
        'Application - Teaching Position',
        'Application - Non-teaching/Teaching-related',
        'COE-Certificate of Employment',
        'Loan Approval and Verification',
        'Leave Application',
        'Retirement',
        'Service Record',
        'Other requests/inquiries',
      ];
    }

    if (key.includes('property and supply')) {
      return [
        'Request/Issuance of Supplies',
        'Other requests/inquiries',
      ];
    }

    if (key.includes('records')) {
      return [
        'CAV-Certification, Authentication, Verification',
        'Receiving & releasing of documents',
        'Certified True Copy (CTC)/Photocopy of documents',
        'Other requests/inquiries',
      ];
    }

    if (key.includes('cid')) {
      return [
        'ALS Enrollment',
        'Borrowing of books/learning materials',
        'Contextualized Learning Resources',
        'Access to LR Portal',
        'Instructional Supervision',
        'Technical assistance',
        'Quality Assurance of Supplementary Learning Resources',
        'Other requests/inquiries',
      ];
    }

    if (key.includes('accounting') || key.includes('budget')) {
      return [
        'Accounting-related',
        'Posting/Updating of Disbursement',
        'ORS-Obligation Request and Status',
        'Other requests/inquiries',
      ];
    }

    if (key.includes('ict')) {
      return [
        'Create/delete/rename/reset user accounts',
        'Troubleshooting of ICT equipment',
        'Uploading of publications',
        'Other requests/inquiries',
      ];
    }

    if (key.includes('legal')) {
      return [
        'Legal advice/opinion',
        'Correction of Entries in School Record',
        'Certificate of No Pending Case',
      ];
    }

    if (key.includes('sgod') && !key.includes('private')) {
      return [
        'EBEIS/LIS/NAT Data and Performance Indicators',
        'Basic Education Data',
        'Private school-related',
        'Other requests/inquiries',
      ];
    }

    if (key.includes('sgod') && key.includes('private')) {
      return [
        'Other private school concerns',
        'Private schools permit/recognition/renewal',
        'Special Orders-graduation of private schools learners',
        'No Increase in tuition/other school fees',
        'Increase in tuition/other school fees (TOSF)',
      ];
    }

    return [];
  }

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
