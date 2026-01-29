const fileInput = document.getElementById('fileInput');
const statusEl = document.getElementById('status');
const tableWrapper = document.getElementById('tableWrapper');
const table = document.getElementById('excelTable');
const thead = table.querySelector('thead');
const tbody = table.querySelector('tbody');
const headersListContainer = document.getElementById('headersListContainer');
const headersList = document.getElementById('headersList');

// Modal preview elements
const previewModalBtn = document.getElementById('previewModalBtn');
const previewHeadersBtn = document.getElementById('previewHeadersBtn');
const previewModal = document.getElementById('previewModal');
const closePreviewModalBtn = document.getElementById('closePreviewModal');
const modalTable = document.getElementById('excelTableModal');
const modalThead = modalTable.querySelector('thead');
const modalTbody = modalTable.querySelector('tbody');

// Headers preview modal elements
const headersPreviewModal = document.getElementById('headersPreviewModal');
const closeHeadersModalBtn = document.getElementById('closeHeadersModal');

// "View all reports" button (generate for all offices at once)
const viewAllReportsBtn = document.getElementById('viewAllReportsBtn');

// Containers for distinct "Service availed - ..." values per office
const servicesListContainerSDS = document.getElementById('servicesListContainerSDS');
const servicesListSDS = document.getElementById('servicesListSDS');
const servicesListContainerASDS = document.getElementById('servicesListContainerASDS');
const servicesListASDS = document.getElementById('servicesListASDS');
const servicesListContainerCashGSP = document.getElementById('servicesListContainerCashGSP');
const servicesListCashGSP = document.getElementById('servicesListCashGSP');
const servicesListContainerPersonnel = document.getElementById('servicesListContainerPersonnel');
const servicesListPersonnel = document.getElementById('servicesListPersonnel');
const servicesListContainerPropertySupply = document.getElementById('servicesListContainerPropertySupply');
const servicesListPropertySupply = document.getElementById('servicesListPropertySupply');
const servicesListContainerRecords = document.getElementById('servicesListContainerRecords');
const servicesListRecords = document.getElementById('servicesListRecords');
const servicesListContainerCID = document.getElementById('servicesListContainerCID');
const servicesListCID = document.getElementById('servicesListCID');
const servicesListContainerAccountingBudget = document.getElementById('servicesListContainerAccountingBudget');
const servicesListAccountingBudget = document.getElementById('servicesListAccountingBudget');
const servicesListContainerICT = document.getElementById('servicesListContainerICT');
const servicesListICT = document.getElementById('servicesListICT');
const servicesListContainerLegal = document.getElementById('servicesListContainerLegal');
const servicesListLegal = document.getElementById('servicesListLegal');
const servicesListContainerSGOD = document.getElementById('servicesListContainerSGOD');
const servicesListSGOD = document.getElementById('servicesListSGOD');
const servicesListContainerSGODPrivate = document.getElementById('servicesListContainerSGODPrivate');
const servicesListSGODPrivate = document.getElementById('servicesListSGODPrivate');

// Keep the latest parsed sheet in memory so we can compute report totals later
let currentHeaderRow = null;
let currentBodyRows = null;

fileInput.addEventListener('change', handleFile, false);
if (previewModalBtn && previewModal && closePreviewModalBtn) {
   previewModalBtn.addEventListener('click', openPreviewModal);
   closePreviewModalBtn.addEventListener('click', closePreviewModal);
}
if (previewHeadersBtn && headersPreviewModal && closeHeadersModalBtn) {
   previewHeadersBtn.addEventListener('click', openHeadersModal);
   closeHeadersModalBtn.addEventListener('click', closeHeadersModal);
}

setupOfficeReportButtons();
restoreLastSheetFromStorage();
setupViewAllReportsButton();

function handleFile(event) {
   const file = event.target.files[0];
   if (!file) {
      statusEl.textContent = 'No file selected.';
      tableWrapper.classList.add('hidden');
      headersListContainer.classList.add('hidden');
      hideAllServiceContainers();
      if (previewModalBtn) {
         previewModalBtn.classList.add('hidden');
      }
      if (previewHeadersBtn) {
         previewHeadersBtn.classList.add('hidden');
      }
      if (headersPreviewModal) {
         headersPreviewModal.classList.add('hidden');
      }
      return;
   }

   statusEl.textContent = `Reading file: ${file.name} ...`;

   const reader = new FileReader();

   reader.onload = function (e) {
      const data = new Uint8Array(e.target.result);
      try {
         const workbook = XLSX.read(data, { type: 'array' });

         // Get the first sheet name
         const firstSheetName = workbook.SheetNames[0];
         const worksheet = workbook.Sheets[firstSheetName];

         // Convert sheet to JSON array of arrays
         // header: 1  -> return rows as simple arrays
         // defval: '' -> use empty string for empty cells
         // raw: false -> use formatted text (e.g., proper dates/times) instead of raw numbers
         const sheetData = XLSX.utils.sheet_to_json(worksheet, {
            header: 1,
            defval: '',
            raw: false,
         });

         if (!sheetData || sheetData.length === 0) {
            statusEl.textContent = 'Sheet is empty.';
            tableWrapper.classList.add('hidden');
            headersListContainer.classList.add('hidden');
            hideAllServiceContainers();
            return;
         }

         renderTableAndLists(sheetData);

         // Persist raw sheet data so it survives a page refresh
         try {
            localStorage.setItem('sdoLastSheetData', JSON.stringify(sheetData));
         } catch (e) {
            console.warn('Could not store sheet data in localStorage', e);
         }
         statusEl.textContent = `Loaded sheet: ${firstSheetName} (${sheetData.length - 1} data rows)`;
      } catch (err) {
         console.error(err);
         statusEl.textContent = 'Error reading file. Check console for details.';
         tableWrapper.classList.add('hidden');
         headersListContainer.classList.add('hidden');
         hideAllServiceContainers();
         if (previewModalBtn) {
            previewModalBtn.classList.add('hidden');
         }
         if (previewHeadersBtn) {
            previewHeadersBtn.classList.add('hidden');
         }
         if (headersPreviewModal) {
            headersPreviewModal.classList.add('hidden');
         }
      }
   };

   reader.onerror = function () {
      statusEl.textContent = 'Failed to read file.';
      tableWrapper.classList.add('hidden');
      headersListContainer.classList.add('hidden');
      hideAllServiceContainers();
      if (previewModalBtn) {
         previewModalBtn.classList.add('hidden');
      }
      if (previewHeadersBtn) {
         previewHeadersBtn.classList.add('hidden');
      }
   };

   reader.readAsArrayBuffer(file);
}

function renderTableAndLists(sheetData) {
   // Clear any previous content
   thead.innerHTML = '';
   tbody.innerHTML = '';
   headersList.innerHTML = '';

   const headerRow = sheetData[0] || [];
   const bodyRows = sheetData.slice(1);

   // Store globally for report computations
   currentHeaderRow = headerRow;
   currentBodyRows = bodyRows;

   // Render table header
   const trHead = document.createElement('tr');
   headerRow.forEach((cellValue, index) => {
      const th = document.createElement('th');
      th.textContent = cellValue !== '' ? cellValue : `Column ${index + 1}`;
      trHead.appendChild(th);
   });
   thead.appendChild(trHead);

   // Render vertical list of headers
   headerRow.forEach((cellValue, index) => {
      const li = document.createElement('li');
      li.textContent = cellValue !== '' ? cellValue : `Column ${index + 1}`;
      headersList.appendChild(li);
   });

   // Render body rows
   bodyRows.forEach((row) => {
      const tr = document.createElement('tr');
      headerRow.forEach((_, colIndex) => {
         const td = document.createElement('td');
         td.textContent = row[colIndex] !== undefined ? row[colIndex] : '';
         tr.appendChild(td);
      });
      tbody.appendChild(tr);
   });

   // Render distinct services from "Service availed - SDS" column (if present)
   renderDistinctServices(headerRow, bodyRows);

   if (previewModalBtn) {
      previewModalBtn.classList.remove('hidden');
   }

   if (previewHeadersBtn) {
      previewHeadersBtn.classList.remove('hidden');
   }

   if (viewAllReportsBtn) {
      viewAllReportsBtn.classList.remove('hidden');
   }
}

function restoreLastSheetFromStorage() {
   try {
      const raw = localStorage.getItem('sdoLastSheetData');
      if (!raw) {
         return;
      }

      const sheetData = JSON.parse(raw);
      if (!Array.isArray(sheetData) || sheetData.length === 0) {
         return;
      }

      renderTableAndLists(sheetData);
      statusEl.textContent = `Restored last uploaded sheet (${sheetData.length - 1} data rows)`;
   } catch (e) {
      console.warn('Could not restore sheet data from localStorage', e);
   }
}

function computeReportTotals(office, period) {
   const totals = {
      sexMale: 0,
      sexFemale: 0,
      sexTotal: 0,
      age19Lower: 0,
      age20_34: 0,
      age35_49: 0,
      age50_64: 0,
      ageTotal: 0,
      custBusiness: 0,
      custCitizen: 0,
      custGovernment: 0,
      custTotal: 0,
      // Will be filled below
      serviceRows: [],
      ccTotals: null,
   };

   if (!currentHeaderRow || !currentBodyRows || currentBodyRows.length === 0) {
      alert('No Excel data loaded. Please upload a file first.');
      return totals;
   }

   const headerRow = currentHeaderRow;
   const bodyRows = currentBodyRows;

   // Column indices for common fields
   const sexColIndex = headerRow.findIndex((h) => h && String(h).toLowerCase().includes('sex'));

   // Prefer an exact 'Age' header (trimmed, case-insensitive)
   let ageColIndex = headerRow.findIndex((h) => {
      if (!h) return false;
      return String(h).trim().toLowerCase() === 'age';
   });
   const custColIndex = headerRow.findIndex((h) => {
      if (!h) return false;
      const text = String(h).toLowerCase();
      return text.includes('customer type') || text.includes('client type') || text.includes('type of customer');
   });

   // Fallback: if we couldn't find an Age column by header text but we do have a Sex column,
   // assume Age is the column immediately before or after Sex (as in your sheet screenshot).
   if (ageColIndex === -1 && sexColIndex !== -1) {
      if (sexColIndex > 0) {
         ageColIndex = sexColIndex - 1;
      } else if (sexColIndex < headerRow.length - 1) {
         ageColIndex = sexColIndex + 1;
      }
   }

   // Date column for filtering by period (month)
   // In your sheet this is the "Completion time" column with values like
   // MM/DD/YYYY HH:MM:SS AM/PM. We match that header first, then fall back to any 'date' header.
   let dateColIndex = headerRow.findIndex((h) => {
      if (!h) return false;
      const text = String(h).toLowerCase();
      return text.includes('completion time');
   });
   if (dateColIndex === -1) {
      dateColIndex = headerRow.findIndex((h) => h && String(h).toLowerCase().includes('date'));
   }

   // Office-specific "Service availed - ..." column index
   const officeMatchers = {
      SDS: (text) => text.includes('service availed') && text.includes('sds'),
      ASDS: (text) => text.includes('service availed') && text.includes('asds'),
      'Cash, General Services, Procurement': (text) => text.includes('service availed') && text.includes('cash'),
      Personnel: (text) => text.includes('service availed') && text.includes('personnel'),
      'Property and Supply': (text) => text.includes('service availed') && text.includes('property and supply'),
      Records: (text) => text.includes('service availed') && text.includes('records'),
      CID: (text) => text.includes('service availed') && text.includes('cid'),
      'Accounting, Budget': (text) => text.includes('service availed') && text.includes('accounting'),
      ICT: (text) => text.includes('service availed') && text.includes('ict'),
      Legal: (text) => text.includes('service availed') && text.includes('legal'),
      SGOD: (text) => text.includes('service availed') && !text.includes('sgod (private') && text.includes('sgod'),
      'SGOD (Private school-related)': (text) => text.includes('service availed') && text.includes('sgod (private'),
   };

   let officeColIndex = -1;
   const officeMatcher = officeMatchers[office];
   if (officeMatcher) {
      officeColIndex = headerRow.findIndex((h) => {
         if (!h) return false;
         const text = String(h).toLowerCase();
         return officeMatcher(text);
      });
   }

   function rowMatchesOffice(row) {
      if (officeColIndex === -1) return true; // if we can't find the column, don't filter by office
      const value = row[officeColIndex];
      if (value === undefined || value === null) return false;
      const trimmed = String(value).trim();
      return trimmed !== '';
   }

   function monthFromPeriod(p) {
      if (!p) return null;
      const val = String(p).toLowerCase();
      if (val === 'jan-dec' || val === 'whole year') return null;
      const map = {
         january: 0,
         february: 1,
         march: 2,
         april: 3,
         may: 4,
         june: 5,
         july: 6,
         august: 7,
         september: 8,
         october: 9,
         november: 10,
         december: 11,
      };
      return map[val] ?? null;
   }

   const targetMonth = monthFromPeriod(period);

   // Citizen's Charter columns (office-level, not per service)
   // Match using distinctive parts of the actual question texts from the Excel.
   const cc1ColIndex = headerRow.findIndex((h) => {
      if (!h) return false;
      const text = String(h).toLowerCase();
      // "Are you aware of the Citizen's Charter - document of the SDO services and requirements?"
      return text.includes("are you aware of the citizen's charter") || text.includes('are you aware of the citizen');
   });
   const cc2ColIndex = headerRow.findIndex((h) => {
      if (!h) return false;
      const text = String(h).toLowerCase();
      // "Did you see the SDO Citizen's Charter (online or posted in the office)?"
      return text.includes("did you see the sdo citizen's charter") ||
         (text.includes('did you see') && text.includes('citizen') && text.includes('charter'));
   });
   const cc3ColIndex = headerRow.findIndex((h) => {
      if (!h) return false;
      const text = String(h).toLowerCase();
      // "Did you use the SDO Citizen's Charter as a guide for the service you availed"
      return text.includes("did you use the sdo citizen's charter") ||
         (text.includes('did you use') && text.includes('citizen') && text.includes('charter'));
   });

   let cc1Yes = 0;
   let cc1No = 0;
   let cc2Yes = 0;
   let cc2No = 0;
   let cc3Yes = 0;
   let cc3No = 0;
   let totalRespondents = 0; // rows for this office + period

   // Per-service counts for this office + period
   const serviceCounts = new Map();

   // SQD columns (client satisfaction questions)
   // Match using distinctive parts of the actual question texts.
   function findSqdColIndex(matchFn) {
      return headerRow.findIndex((h) => {
         if (!h) return false;
         const text = String(h).toLowerCase();
         return matchFn(text);
      });
   }

   const sqd1ColIndex = findSqdColIndex((text) =>
      text.includes('sqd1') && text.includes('i spent an acceptable amount of time')
   );
   const sqd2ColIndex = findSqdColIndex((text) =>
      text.includes('sqd2') && text.includes('accurately informed')
   );
   const sqd3ColIndex = findSqdColIndex((text) =>
      text.includes('sqd3') && text.includes('simple and convenient')
   );
   const sqd4ColIndex = findSqdColIndex((text) =>
      (text.includes('sqd4') || text.includes('sdq4')) && text.includes('easily found information')
   );
   const sqd5ColIndex = findSqdColIndex((text) =>
      text.includes('sqd5') && text.includes('paid an acceptable amount of fees')
   );
   const sqd6ColIndex = findSqdColIndex((text) =>
      text.includes('sqd6') && text.includes('confident my transaction was secure')
   );
   const sqd7ColIndex = findSqdColIndex((text) =>
      text.includes('sqd7') && text.includes("support was quick to respond")
   );
   const sqd8ColIndex = findSqdColIndex((text) =>
      text.includes('sqd8') && text.includes('got what i needed')
   );

   const sqdAgg = [
      { label: 'SQD1 (Responsiveness)', sa5: 0, a4: 0, n3: 0, d2: 0, sd1: 0, na: 0, totalRespondents: 0, totalRatedScore: 0 },
      { label: 'SQD2 (Reliability)', sa5: 0, a4: 0, n3: 0, d2: 0, sd1: 0, na: 0, totalRespondents: 0, totalRatedScore: 0 },
      { label: 'SQD3 (Access)', sa5: 0, a4: 0, n3: 0, d2: 0, sd1: 0, na: 0, totalRespondents: 0, totalRatedScore: 0 },
      { label: 'SQD4 (Communication)', sa5: 0, a4: 0, n3: 0, d2: 0, sd1: 0, na: 0, totalRespondents: 0, totalRatedScore: 0 },
      { label: 'SQD5 (Costs)', sa5: 0, a4: 0, n3: 0, d2: 0, sd1: 0, na: 0, totalRespondents: 0, totalRatedScore: 0 },
      { label: 'SQD6 (Integrity)', sa5: 0, a4: 0, n3: 0, d2: 0, sd1: 0, na: 0, totalRespondents: 0, totalRatedScore: 0 },
      { label: 'SQD7 (Assurance)', sa5: 0, a4: 0, n3: 0, d2: 0, sd1: 0, na: 0, totalRespondents: 0, totalRatedScore: 0 },
      { label: 'SQD8 (Outcome)', sa5: 0, a4: 0, n3: 0, d2: 0, sd1: 0, na: 0, totalRespondents: 0, totalRatedScore: 0 },
   ];

   function rowMatchesPeriod(row) {
      if (targetMonth === null || dateColIndex === -1) return true;
      const value = row[dateColIndex];
      if (!value) return false;
      const date = new Date(value);
      if (Number.isNaN(date.getTime())) return false;
      return date.getMonth() === targetMonth;
   }

   bodyRows.forEach((row) => {
      if (!rowMatchesOffice(row) || !rowMatchesPeriod(row)) {
         return;
      }

      // Count respondent for CC totals and service counts
      totalRespondents += 1;

      // Service availed (based on office-specific column)
      if (officeColIndex !== -1) {
         const rawService = row[officeColIndex];
         if (rawService !== undefined && rawService !== null) {
            const serviceName = String(rawService).trim();
            if (serviceName) {
               const current = serviceCounts.get(serviceName) || 0;
               serviceCounts.set(serviceName, current + 1);
            }
         }
      }

      // Sex
      if (sexColIndex !== -1) {
         const rawSex = row[sexColIndex];
         if (rawSex !== undefined && rawSex !== null) {
            const sex = String(rawSex).trim().toLowerCase();
            if (sex === 'male' || sex === 'm') {
               totals.sexMale += 1;
               totals.sexTotal += 1;
            } else if (sex === 'female' || sex === 'f') {
               totals.sexFemale += 1;
               totals.sexTotal += 1;
            }
         }
      }

      // Age
      if (ageColIndex !== -1) {
         const rawAge = row[ageColIndex];
         if (rawAge !== undefined && rawAge !== null && String(rawAge).trim() !== '') {
            const rawStr = String(rawAge).trim();

            // First try numeric age (e.g., 23, 45)
            const num = parseFloat(rawStr.replace(/[^0-9.]/g, ''));
            if (!Number.isNaN(num)) {
               if (num <= 19) {
                  totals.age19Lower += 1;
               } else if (num >= 20 && num <= 34) {
                  totals.age20_34 += 1;
               } else if (num >= 35 && num <= 49) {
                  totals.age35_49 += 1;
               } else if (num >= 50 && num <= 64) {
                  totals.age50_64 += 1;
               }
               totals.ageTotal += 1;
            } else {
               // Fallback: textual age brackets like "19-lower", "20-34", "35-49", "50-64"
               const lower = rawStr.toLowerCase();
               if (lower.includes('19') && lower.includes('lower')) {
                  totals.age19Lower += 1;
                  totals.ageTotal += 1;
               } else if (lower.includes('20-34') || (lower.includes('20') && lower.includes('34'))) {
                  totals.age20_34 += 1;
                  totals.ageTotal += 1;
               } else if (lower.includes('35-49') || (lower.includes('35') && lower.includes('49'))) {
                  totals.age35_49 += 1;
                  totals.ageTotal += 1;
               } else if (lower.includes('50-64') || (lower.includes('50') && lower.includes('64'))) {
                  totals.age50_64 += 1;
                  totals.ageTotal += 1;
               }
            }
         }
      }

      // Customer type
      if (custColIndex !== -1) {
         const rawCust = row[custColIndex];
         if (rawCust !== undefined && rawCust !== null) {
            const text = String(rawCust).trim().toLowerCase();
            if (!text) return;
            if (text.includes('business')) {
               totals.custBusiness += 1;
               totals.custTotal += 1;
            } else if (text.includes('citizen') || text.includes('private')) {
               totals.custCitizen += 1;
               totals.custTotal += 1;
            } else if (text.includes('government') || text.includes('gov')) {
               totals.custGovernment += 1;
               totals.custTotal += 1;
            }
         }
      }

      // Citizen's Charter CC1–CC3 (office-level, per respondent)
      if (cc1ColIndex !== -1) {
         const raw = row[cc1ColIndex];
         const cat = classifyCcResponse(raw);
         if (cat === 'yes') cc1Yes += 1;
         else if (cat === 'no') cc1No += 1;
      }

      if (cc2ColIndex !== -1) {
         const raw = row[cc2ColIndex];
         const cat = classifyCcResponse(raw);
         if (cat === 'yes') cc2Yes += 1;
         else if (cat === 'no') cc2No += 1;
      }

      if (cc3ColIndex !== -1) {
         const raw = row[cc3ColIndex];
         const cat = classifyCcResponse(raw);
         if (cat === 'yes') cc3Yes += 1;
         else if (cat === 'no') cc3No += 1;
      }

      // Client Satisfaction SQD1–SQD8 (per respondent, per office+period)
      const sqdCols = [
         sqd1ColIndex,
         sqd2ColIndex,
         sqd3ColIndex,
         sqd4ColIndex,
         sqd5ColIndex,
         sqd6ColIndex,
         sqd7ColIndex,
         sqd8ColIndex,
      ];

      sqdCols.forEach((colIndex, idx) => {
         if (colIndex === -1) return;
         const raw = row[colIndex];
         const cat = classifySqdResponse(raw);
         const agg = sqdAgg[idx];
         if (!agg || cat === 'none') return;

         if (cat === 'sa5') {
            agg.sa5 += 1;
            agg.totalRespondents += 1;
            agg.totalRatedScore += 5;
         } else if (cat === 'a4') {
            agg.a4 += 1;
            agg.totalRespondents += 1;
            agg.totalRatedScore += 4;
         } else if (cat === 'n3') {
            agg.n3 += 1;
            agg.totalRespondents += 1;
            agg.totalRatedScore += 3;
         } else if (cat === 'd2') {
            agg.d2 += 1;
            agg.totalRespondents += 1;
            agg.totalRatedScore += 2;
         } else if (cat === 'sd1') {
            agg.sd1 += 1;
            agg.totalRespondents += 1;
            agg.totalRatedScore += 1;
         } else if (cat === 'na') {
            agg.na += 1;
         }
      });
   });

   // Build service rows array (for PDF table)
   totals.serviceRows = Array.from(serviceCounts.entries()).map(([name, count]) => ({
      name,
      count,
   }));

   // Compute Did Not Specify using: totalRespondents - (yes + no)
   const cc1DidNotSpecify = Math.max(0, totalRespondents - (cc1Yes + cc1No));
   const cc2DidNotSpecify = Math.max(0, totalRespondents - (cc2Yes + cc2No));
   const cc3DidNotSpecify = Math.max(0, totalRespondents - (cc3Yes + cc3No));

   totals.ccTotals = {
      totalRespondents,
      cc1: { yes: cc1Yes, no: cc1No, didNotSpecify: cc1DidNotSpecify },
      cc2: { yes: cc2Yes, no: cc2No, didNotSpecify: cc2DidNotSpecify },
      cc3: { yes: cc3Yes, no: cc3No, didNotSpecify: cc3DidNotSpecify },
   };

   // Build client satisfaction rows (SQD1–SQD8) for the third table
   totals.clientSatisfactionRows = sqdAgg.map((agg) => {
      const avg = agg.totalRespondents > 0
         ? (agg.totalRatedScore / agg.totalRespondents)
         : 0;
      return {
         label: agg.label,
         sa5: agg.sa5,
         a4: agg.a4,
         n3: agg.n3,
         d2: agg.d2,
         sd1: agg.sd1,
         na: agg.na,
         totalRespondents: agg.totalRespondents,
         totalRatedScore: agg.totalRatedScore,
         // Format to 2 decimal places like in your sample table
         averageScore: avg ? avg.toFixed(2) : '',
      };
   });

   return totals;
}

function classifyCcResponse(raw) {
   if (raw === undefined || raw === null) return 'none';
   const text = String(raw).trim().toLowerCase();
   if (!text) return 'none';

   // Only look at the first few characters so variants like
   // "Yes - but it was hard to find" and "Yes - it was easy to find" are both yes.
   const first3 = text.slice(0, 3); // e.g., 'yes', 'no ', 'ye-'
   const first2 = text.slice(0, 2); // e.g., 'no', 'ye'

   if (first3 === 'yes') {
      return 'yes';
   }
   if (first2 === 'no') {
      return 'no';
   }

   // Fallback to some older patterns just in case
   if (text === 'y' || text === '1' || text.includes('very satisfied')) {
      return 'yes';
   }
   if (text === 'n' || text === '0' || text.includes('not satisfied')) {
      return 'no';
   }
   return 'none';
}

function classifySqdResponse(raw) {
   if (raw === undefined || raw === null) return 'none';
   const text = String(raw).trim().toLowerCase();
   if (!text) return 'none';

   // The sheet uses strings like "Strongly Agree (5)", "Agree (4)", "Not applicable".
   if (text.startsWith('strongly agree')) return 'sa5';
   if (text.startsWith('agree')) return 'a4';
   if (text.startsWith('neither agree') || text.startsWith('neither agree nor disagree')) return 'n3';
   if (text.startsWith('disagree')) return 'd2';
   if (text.startsWith('strongly disagree')) return 'sd1';
   if (text.startsWith('not applicable')) return 'na';

   return 'none';
}

function renderDistinctServices(headerRow, bodyRows) {
   const configs = [
      {
         container: servicesListContainerSDS,
         list: servicesListSDS,
         match: (text) => text.includes('service availed') && text.includes('sds'),
      },
      {
         container: servicesListContainerASDS,
         list: servicesListASDS,
         match: (text) => text.includes('service availed') && text.includes('asds'),
      },
      {
         container: servicesListContainerCashGSP,
         list: servicesListCashGSP,
         match: (text) => text.includes('service availed') && text.includes('cash'),
      },
      {
         container: servicesListContainerPersonnel,
         list: servicesListPersonnel,
         match: (text) => text.includes('service availed') && text.includes('personnel'),
      },
      {
         container: servicesListContainerPropertySupply,
         list: servicesListPropertySupply,
         match: (text) => text.includes('service availed') && text.includes('property and supply'),
      },
      {
         container: servicesListContainerRecords,
         list: servicesListRecords,
         match: (text) => text.includes('service availed') && text.includes('records'),
      },
      {
         container: servicesListContainerCID,
         list: servicesListCID,
         match: (text) => text.includes('service availed') && text.includes('cid'),
      },
      {
         container: servicesListContainerAccountingBudget,
         list: servicesListAccountingBudget,
         match: (text) => text.includes('service availed') && text.includes('accounting'),
      },
      {
         container: servicesListContainerICT,
         list: servicesListICT,
         match: (text) => text.includes('service availed') && text.includes('ict'),
      },
      {
         container: servicesListContainerLegal,
         list: servicesListLegal,
         match: (text) => text.includes('service availed') && text.includes('legal'),
      },
      {
         container: servicesListContainerSGOD,
         list: servicesListSGOD,
         match: (text) => text.includes('service availed') && text.includes('sgod (') === false && text.includes('sgod'),
      },
      {
         container: servicesListContainerSGODPrivate,
         list: servicesListSGODPrivate,
         match: (text) => text.includes('service availed') && text.includes('sgod (private'),
      },
   ];

   configs.forEach(({ container, list, match }) => {
      if (!container || !list) {
         return;
      }

      // Find matching column index
      const colIndex = headerRow.findIndex((headerCell) => {
         if (!headerCell) return false;
         const text = String(headerCell).toLowerCase();
         return match(text);
      });

      if (colIndex === -1) {
         container.classList.add('hidden');
         list.innerHTML = '';
         return;
      }

      const uniqueServices = new Set();

      bodyRows.forEach((row) => {
         const value = row[colIndex];
         if (value !== undefined && value !== null) {
            const trimmed = String(value).trim();
            if (trimmed) {
               uniqueServices.add(trimmed);
            }
         }
      });

      list.innerHTML = '';

      if (uniqueServices.size === 0) {
         const li = document.createElement('li');
         li.textContent = 'No service values found in this sheet.';
         list.appendChild(li);
      } else {
         uniqueServices.forEach((service) => {
            const li = document.createElement('li');
            li.textContent = service;
            list.appendChild(li);
         });
      }

      container.classList.remove('hidden');
   });
}

function setupViewAllReportsButton() {
   if (!viewAllReportsBtn) return;

   viewAllReportsBtn.addEventListener('click', () => {
      if (typeof window.generatePaginatedPdf !== 'function') {
         alert('PDF generator is not loaded.');
         return;
      }

      // Use "Jan-Dec" (whole year) for all offices, no month filter
      const period = 'Jan-Dec';

      const sections = document.querySelectorAll('.office-report-section');
      if (!sections.length) {
         alert('No offices found to generate reports for.');
         return;
      }

      const multiOffices = [];

      sections.forEach((section) => {
         const office = section.getAttribute('data-office') || '';
         if (!office) return;

         const totals = computeReportTotals(office, period);
         multiOffices.push({ office, period, totals });
      });

      if (!multiOffices.length) {
         alert('No valid offices found to generate reports for.');
         return;
      }

      window.generatePaginatedPdf({
         preview: true,
         multiOffices,
      });
   });
}

function hideAllServiceContainers() {
   const containers = [
      servicesListContainerSDS,
      servicesListContainerASDS,
      servicesListContainerCashGSP,
      servicesListContainerPersonnel,
      servicesListContainerPropertySupply,
      servicesListContainerRecords,
      servicesListContainerCID,
      servicesListContainerAccountingBudget,
      servicesListContainerICT,
      servicesListContainerLegal,
      servicesListContainerSGOD,
      servicesListContainerSGODPrivate,
   ];

   containers.forEach((c) => {
      if (c) {
         c.classList.add('hidden');
      }
   });
}

function setupOfficeReportButtons() {
   const sections = document.querySelectorAll('.office-report-section');
   sections.forEach((section) => {
      const button = section.querySelector('.view-report-button');
      const select = section.querySelector('.rating-period');
      if (!button || !select) return;

      const office = section.getAttribute('data-office') || '';

      button.addEventListener('click', () => {
         const period = select.value;
         if (typeof window.generatePaginatedPdf !== 'function') {
            alert('PDF generator is not loaded.');
            return;
         }

         const totals = computeReportTotals(office, period);

         window.generatePaginatedPdf({
            preview: true,
            office,
            period,
            totals,
         });
      });
   });
}

function toggleHeadersPreview() {
   if (!headersListContainer) return;
   headersListContainer.classList.toggle('hidden');
}

function openPreviewModal() {
   if (!previewModal) return;

   // Clone current table content into modal table
   modalThead.innerHTML = thead.innerHTML;
   modalTbody.innerHTML = tbody.innerHTML;

   previewModal.classList.remove('hidden');
}

function closePreviewModal() {
   if (!previewModal) return;
   previewModal.classList.add('hidden');
}

function openHeadersModal() {
   if (!headersPreviewModal) return;
   // headersListContainer already contains the latest headers; just show the modal
   headersPreviewModal.classList.remove('hidden');
}

function closeHeadersModal() {
   if (!headersPreviewModal) return;
   headersPreviewModal.classList.add('hidden');
}

