import React, { useEffect, useState } from 'react';
import { 
  Box, 
  Typography, 
  Paper, 
  Table, 
  TableBody, 
  TableCell, 
  TableContainer, 
  TableHead, 
  TableRow, 
  Button,
  CircularProgress,
  Tabs,
  Tab,
  Autocomplete,
  TextField,
  Grid,
  TablePagination,
  Menu,
  MenuItem,
  Divider
} from '@mui/material';
import { Download, Clear, FileDownload } from '@mui/icons-material';
import { getCollectionEntries, getSignatories, getRPTCollections, type CollectionEntry } from '../services/googleSheets';
import type { Signatory, RPTCollectionItem } from '../types/rcd';
import * as XLSX from 'xlsx';

const getBookletRange = (num: number) => {
  const lastTwo = num % 100;
  if (lastTwo >= 1 && lastTwo <= 50) {
    const start = Math.floor(num / 100) * 100 + 1;
    const end = start + 49;
    return { start, end };
  } else {
    let start = Math.floor(num / 100) * 100 + 51;
    if (lastTwo === 0) {
      start = num - 49;
    }
    const end = start + 49;
    return { start, end };
  }
};

interface BookletAccountabilityEntry {
  begQty: number;
  begFrom: string;
  begTo: string;
  issQty: number;
  issFrom: string;
  issTo: string;
  endQty?: number;
  endFrom: string;
  endTo: string;
}

const getBookletAccountability = (startOr: string, endOr: string, padLen = 7): BookletAccountabilityEntry[] => {
  const formatOr = (num: number) => num.toString().padStart(padLen, '0');
  const startNum = parseInt(startOr, 10);
  const endNum = parseInt(endOr, 10);
  if (isNaN(startNum) || isNaN(endNum) || startNum > endNum) {
    return [];
  }

  const entries: BookletAccountabilityEntry[] = [];
  let currentNum = startNum;

  while (currentNum <= endNum) {
    const booklet = getBookletRange(currentNum);
    
    const issStart = currentNum;
    const issEnd = Math.min(endNum, booklet.end);

    const begStart = issStart;
    const begEnd = booklet.end;
    const begQty = begEnd - begStart + 1;

    const issQty = issEnd - issStart + 1;

    const endStart = issEnd + 1;
    const endEnd = booklet.end;
    const endQty = endEnd - endStart + 1;

    entries.push({
      begQty,
      begFrom: formatOr(begStart),
      begTo: formatOr(begEnd),
      issQty,
      issFrom: formatOr(issStart),
      issTo: formatOr(issEnd),
      endQty: endQty > 0 ? endQty : undefined,
      endFrom: endQty > 0 ? formatOr(endStart) : '',
      endTo: endQty > 0 ? formatOr(endEnd) : ''
    });

    currentNum = booklet.end + 1;
  }

  return entries;
};

export const ReportsPage: React.FC = () => {
  const [collections, setCollections] = useState<CollectionEntry[]>([]);
  const [rptCollections, setRptCollections] = useState<RPTCollectionItem[]>([]);
  const [filteredCollections, setFilteredCollections] = useState<CollectionEntry[]>([]);
  const [filteredRptCollections, setFilteredRptCollections] = useState<RPTCollectionItem[]>([]);
  const [signatories, setSignatories] = useState<Signatory[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [tabValue, setTabValue] = useState(0);

  // Pagination
  const [page, setPage] = useState(0);
  const [rowsPerPage, setRowsPerPage] = useState(10);

  // Filters
  const [selectedAfNos, setSelectedAfNos] = useState<string[]>([]);
  const [selectedSubCategories, setSelectedSubCategories] = useState<string[]>([]);
  const [selectedMainCategories, setSelectedMainCategories] = useState<string[]>([]);
  const [startDate, setStartDate] = useState<string>('');
  const [endDate, setEndDate] = useState<string>('');
  
  // RPT Filters
  const [rptFilterAf56Id, setRptFilterAf56Id] = useState<string | null>(null);
  const [rptStartDate, setRptStartDate] = useState<string>('');
  const [rptEndDate, setRptEndDate] = useState<string>('');
  
  // RPT OR Filters
  const [rptStartOr1, setRptStartOr1] = useState<string | null>(null);
  const [rptEndOr1, setRptEndOr1] = useState<string | null>(null);
  const [rptStartOr2, setRptStartOr2] = useState<string | null>(null);
  const [rptEndOr2, setRptEndOr2] = useState<string | null>(null);

  const [startOr1, setStartOr1] = useState<string | null>(null);
  const [endOr1, setEndOr1] = useState<string | null>(null);
  const [startOr2, setStartOr2] = useState<string | null>(null);
  const [endOr2, setEndOr2] = useState<string | null>(null);

  // Menu State for RPT Print Report
  const [rptReportMenuAnchor, setRptReportMenuAnchor] = useState<null | HTMLElement>(null);

  useEffect(() => {
    const fetchData = async () => {
      try {
        const [collectionsData, signatoriesData, rptData] = await Promise.all([
          getCollectionEntries(),
          getSignatories(),
          getRPTCollections()
        ]);
        setCollections(collectionsData);
        setFilteredCollections(collectionsData);
        setSignatories(signatoriesData);
        setRptCollections(rptData);
        setFilteredRptCollections(rptData);
      } catch (error) {
        console.error('Failed to fetch data', error);
      } finally {
        setIsLoading(false);
      }
    };

    fetchData();
  }, []);

  // Filter Logic
  useEffect(() => {
    let result = collections;

    if (selectedAfNos.length > 0) {
      result = result.filter(item => selectedAfNos.includes(item.afNo));
    }

    if (selectedSubCategories.length > 0) {
      result = result.filter(item => selectedSubCategories.includes(item.subCategory));
    }

    if (selectedMainCategories.length > 0) {
      result = result.filter(item => selectedMainCategories.includes(item.mainCategory));
    }

    if (startDate) {
      result = result.filter(item => {
        const itemDate = new Date(item.date);
        const start = new Date(startDate);
        return itemDate >= start;
      });
    }

    if (endDate) {
      result = result.filter(item => {
        const itemDate = new Date(item.date);
        const end = new Date(endDate);
        return itemDate <= end;
      });
    }

    if ((startOr1 && endOr1) || (startOr2 && endOr2)) {
      result = result.filter(item => {
        if (!item.orNo) return false;
        
        const checkRange = (startStr: string, endStr: string) => {
          const itemOr = parseInt(item.orNo, 10);
          const start = parseInt(startStr, 10);
          const end = parseInt(endStr, 10);
          
          if (!isNaN(itemOr) && !isNaN(start) && !isNaN(end)) {
            return itemOr >= start && itemOr <= end;
          }
          return item.orNo >= startStr && item.orNo <= endStr;
        };

        const inRange1 = (startOr1 && endOr1) ? checkRange(startOr1, endOr1) : false;
        const inRange2 = (startOr2 && endOr2) ? checkRange(startOr2, endOr2) : false;
        
        return inRange1 || inRange2;
      });
    }

    setFilteredCollections(result);
    setPage(0);
  }, [collections, selectedAfNos, selectedSubCategories, selectedMainCategories, startDate, endDate, startOr1, endOr1, startOr2, endOr2]);

  const uniqueAfNos = Array.from(new Set(collections.map(c => c.afNo).filter(Boolean))).sort();
  const uniqueSubCategories = Array.from(new Set(collections.map(c => c.subCategory).filter(Boolean))).sort();
  const uniqueMainCategories = Array.from(new Set(collections.map(c => c.mainCategory).filter(Boolean))).sort();

  // RPT Filter Logic
  useEffect(() => {
    let result = rptCollections;

    if (rptFilterAf56Id) {
      result = result.filter(item => item.af56Id === rptFilterAf56Id);
    }

    if (rptStartDate) {
      result = result.filter(item => item.date >= rptStartDate);
    }

    if (rptEndDate) {
      result = result.filter(item => item.date <= rptEndDate);
    }

    setFilteredRptCollections(result);
    // Only reset page if we are on the RPT tab
    if (tabValue === 2) {
      setPage(0);
    }
  }, [rptCollections, rptFilterAf56Id, rptStartDate, rptEndDate, tabValue]);

  const uniqueRptAf56Ids = Array.from(new Set(rptCollections.map(c => c.af56Id).filter(Boolean))).sort();

  // Calculate valid RPT ORs when AF56 ID is selected
  const validRptOrs = React.useMemo(() => {
    if (!rptFilterAf56Id) return [];
    
    const ors = rptCollections
      .filter(c => c.af56Id === rptFilterAf56Id && c.orNumber)
      .map(c => c.orNumber)
      .filter(Boolean);
      
    return Array.from(new Set(ors)).sort((a, b) => {
      // Try numerical sort
      const numA = parseInt(a, 10);
      const numB = parseInt(b, 10);
      if (!isNaN(numA) && !isNaN(numB)) return numA - numB;
      return a.localeCompare(b);
    });
  }, [rptCollections, rptFilterAf56Id]);

  // Reset RPT OR selection when AF56 ID changes
  useEffect(() => {
    setRptStartOr1(null);
    setRptEndOr1(null);
    setRptStartOr2(null);
    setRptEndOr2(null);
  }, [rptFilterAf56Id]);

  // Automatically set RPT Start OR 2 to the next available OR after End OR 1
  useEffect(() => {
    if (rptEndOr1 && validRptOrs.length > 0) {
      const currentIndex = validRptOrs.indexOf(rptEndOr1);
      if (currentIndex !== -1 && currentIndex < validRptOrs.length - 1) {
        setRptStartOr2(validRptOrs[currentIndex + 1]);
      } else if (currentIndex === validRptOrs.length - 1) {
        setRptStartOr2(null);
      }
    }
  }, [rptEndOr1, validRptOrs]);

  // Calculate valid ORs when exactly one AF No. is selected
  const validOrs = React.useMemo(() => {
    if (selectedAfNos.length !== 1) return [];
    
    const afNo = selectedAfNos[0];
    const ors = collections
      .filter(c => c.afNo === afNo && c.orNo)
      .map(c => c.orNo)
      .filter(Boolean);
      
    return Array.from(new Set(ors)).sort((a, b) => {
      // Try numerical sort
      const numA = parseInt(a, 10);
      const numB = parseInt(b, 10);
      if (!isNaN(numA) && !isNaN(numB)) return numA - numB;
      return a.localeCompare(b);
    });
  }, [collections, selectedAfNos]);

  // Reset OR selection when AF selection changes (if no longer single AF or different AF)
  useEffect(() => {
    if (selectedAfNos.length !== 1) {
      setStartOr1(null);
      setEndOr1(null);
      setStartOr2(null);
      setEndOr2(null);
    }
  }, [selectedAfNos]);

  // Automatically set Start OR 2 to the next available OR after End OR 1
  useEffect(() => {
    if (endOr1 && validOrs.length > 0) {
      const currentIndex = validOrs.indexOf(endOr1);
      if (currentIndex !== -1 && currentIndex < validOrs.length - 1) {
        setStartOr2(validOrs[currentIndex + 1]);
      } else if (currentIndex === validOrs.length - 1) {
        setStartOr2(null);
      }
    }
  }, [endOr1, validOrs]);

  const totalFilteredAmount = filteredCollections.reduce((sum, item) => sum + (item.amount || 0), 0);

  // Calculate Summaries for RCD Summaries Tab
  const afNoSummary = React.useMemo(() => {
    const summary: { [key: string]: number } = {};
    collections.forEach(item => {
      const afNo = item.afNo || 'Unspecified';
      summary[afNo] = (summary[afNo] || 0) + (item.amount || 0);
    });
    return Object.entries(summary)
      .map(([afNo, amount]) => ({ afNo, amount }))
      .sort((a, b) => a.afNo.localeCompare(b.afNo));
  }, [collections]);

  const monthlySummary = React.useMemo(() => {
    const summary: { [key: string]: number } = {};
    collections.forEach(item => {
      if (!item.date) return;
      const date = new Date(item.date);
      if (isNaN(date.getTime())) return;
      
      const monthYear = date.toLocaleDateString('en-PH', { month: 'long', year: 'numeric' });
      summary[monthYear] = (summary[monthYear] || 0) + (item.amount || 0);
    });
    
    // Sort by date (parsing monthYear back to date for sorting)
    return Object.entries(summary)
      .map(([monthYear, amount]) => ({ monthYear, amount }))
      .sort((a, b) => {
        const dateA = new Date(a.monthYear);
        const dateB = new Date(b.monthYear);
        return dateB.getTime() - dateA.getTime(); // Descending order
      });
  }, [collections]);

  const subCategorySummary = React.useMemo(() => {
    const summary: { [key: string]: number } = {};
    collections.forEach(item => {
      const subCat = item.subCategory || 'Unspecified';
      summary[subCat] = (summary[subCat] || 0) + (item.amount || 0);
    });
    return Object.entries(summary)
      .map(([subCategory, amount]) => ({ subCategory, amount }))
      .sort((a, b) => a.subCategory.localeCompare(b.subCategory));
  }, [collections]);

  const mainCategorySummary = React.useMemo(() => {
    const summary: { [key: string]: number } = {};
    collections.forEach(item => {
      const mainCat = item.mainCategory || 'Unspecified';
      summary[mainCat] = (summary[mainCat] || 0) + (item.amount || 0);
    });
    return Object.entries(summary)
      .map(([mainCategory, amount]) => ({ mainCategory, amount }))
      .sort((a, b) => a.mainCategory.localeCompare(b.mainCategory));
  }, [collections]);

  const handleTabChange = (_event: React.SyntheticEvent, newValue: number) => {
    setTabValue(newValue);
    setPage(0);
  };

  const handleChangePage = (_event: unknown, newPage: number) => {
    setPage(newPage);
  };

  const handleChangeRowsPerPage = (event: React.ChangeEvent<HTMLInputElement>) => {
    setRowsPerPage(parseInt(event.target.value, 10));
    setPage(0);
  };

  const handlePrintCover = () => {
    if (selectedAfNos.length !== 1) return;
    const afNo = selectedAfNos[0];

    const filterByRange = (sOr: string | null, eOr: string | null) => {
      if (!sOr || !eOr) return [];
      return collections.filter(item => {
        if (item.afNo !== afNo || !item.orNo) return false;
        
        const itemOr = parseInt(item.orNo, 10);
        const start = parseInt(sOr, 10);
        const end = parseInt(eOr, 10);
        
        if (!isNaN(itemOr) && !isNaN(start) && !isNaN(end)) {
          return itemOr >= start && itemOr <= end;
        }
        return item.orNo >= sOr && item.orNo <= eOr;
      });
    };

    const range1Data = filterByRange(startOr1, endOr1);
    const range2Data = filterByRange(startOr2, endOr2);

    const calculateSummary = (data: CollectionEntry[]) => {
      const summary: { [key: string]: number } = {};
      let total = 0;
      data.forEach(item => {
        const subCat = item.subCategory || 'Unspecified';
        summary[subCat] = (summary[subCat] || 0) + (item.amount || 0);
        total += (item.amount || 0);
      });
      return {
        summary: Object.entries(summary)
          .map(([subCategory, amount]) => ({ subCategory, amount }))
          .filter(item => item.amount !== 0)
          .sort((a, b) => a.subCategory.localeCompare(b.subCategory)),
        total
      };
    };

    const r1 = calculateSummary(range1Data);
    const r2 = calculateSummary(range2Data);

    const getDstItems = (data: CollectionEntry[]) => {
      return data
        .filter(item => item.subCategory === 'DST')
        .sort((a, b) => {
          const orA = parseInt(a.orNo || '0', 10);
          const orB = parseInt(b.orNo || '0', 10);
          return orA - orB;
        });
    };

    const r1DstItems = getDstItems(range1Data);
    const r2DstItems = getDstItems(range2Data);

    const printContent = (_rangeLabel: string, start: string | null, end: string | null, data: { summary: { subCategory: string; amount: number }[]; total: number }, dstItems: CollectionEntry[]) => {
      const dstSection = dstItems.length > 0 ? `
        <div class="section">
          <div class="header-box">DST Details</div>
          <table class="summary-table">
            <thead>
              <tr>
                <th style="text-align: left; border-bottom: 1px solid #eee;">Date</th>
                <th style="text-align: left; border-bottom: 1px solid #eee;">OR Number</th>
                <th style="text-align: right; border-bottom: 1px solid #eee;">Amount</th>
              </tr>
            </thead>
            <tbody>
              ${dstItems.map(item => `
                <tr>
                  <td>${item.date ? new Date(item.date).toLocaleDateString('en-PH') : '-'}</td>
                  <td>${item.orNo}</td>
                  <td class="text-right">₱ ${(item.amount || 0).toLocaleString('en-PH', { minimumFractionDigits: 2 })}</td>
                </tr>
              `).join('')}
              <tr class="total-row">
                <td colspan="2">Total</td>
                <td class="text-right">₱ ${dstItems.reduce((sum, item) => sum + (item.amount || 0), 0).toLocaleString('en-PH', { minimumFractionDigits: 2 })}</td>
              </tr>
            </tbody>
          </table>
        </div>
      ` : '';

      return `
      <div class="column">
        <div class="section center-text">
          <div class="label">OR Numbers</div>
          <div class="value">${start || 'N/A'} - ${end || 'N/A'}</div>
        </div>
        
        <div class="section center-text">
          <div class="label">AMOUNT</div>
          <div class="value">₱ ${data.total.toLocaleString('en-PH', { minimumFractionDigits: 2 })}</div>
        </div>

        <div class="section">
          <div class="header-box">GENERAL FUND</div>
          <table class="summary-table">
            <tbody>
              ${data.summary.map(item => `
                <tr>
                  <td>${item.subCategory}</td>
                  <td class="text-right">₱ ${item.amount.toLocaleString('en-PH', { minimumFractionDigits: 2 })}</td>
                </tr>
              `).join('')}
            </tbody>
          </table>
        </div>
        ${dstSection}
      </div>
    `;
    };

    const printWindow = window.open('', '_blank');
    if (printWindow) {
      printWindow.document.write(`
        <html>
          <head>
            <title>Print Cover - AF No. ${afNo}</title>
            <style>
              @page { size: Letter portrait; margin: 0.5in; }
              body { font-family: Arial, sans-serif; margin: 0; padding: 0; -webkit-print-color-adjust: exact; }
              .container { display: flex; width: 100%; height: 100vh; }
              .column { flex: 1; padding: 0 20px; box-sizing: border-box; }
              .column:first-child { border-right: 1px dashed #ccc; }
              .section { margin-bottom: 30px; }
              .center-text { text-align: center; }
              .label { font-size: 14px; font-weight: bold; color: #666; text-transform: uppercase; margin-bottom: 5px; }
              .value { font-size: 18px; font-weight: bold; text-decoration: underline; }
              .header-box { 
                background-color: #f0f0f0; 
                padding: 10px; 
                text-align: center; 
                font-weight: bold; 
                border: 1px solid #000;
                margin-bottom: 15px;
              }
              .summary-table { width: 100%; border-collapse: collapse; font-size: 14px; }
              .summary-table td, .summary-table th { padding: 5px 0; border-bottom: 1px solid #eee; }
              .summary-table .text-right { text-align: right; }
              .summary-table .total-row td { font-weight: bold; border-top: 1px solid #000; border-bottom: none; padding-top: 10px; }
            </style>
          </head>
          <body>
            <div class="container">
              ${printContent('Range 1', startOr1, endOr1, r1, r1DstItems)}
              ${printContent('Range 2', startOr2, endOr2, r2, r2DstItems)}
            </div>
          </body>
        </html>
      `);
      printWindow.document.close();
      printWindow.focus();
      printWindow.print();
    }
  };

  const handlePrintReport = () => {
    // 1. Prepare Data
    const reportData = filteredCollections;
    const totalAmount = reportData.reduce((sum, item) => sum + (item.amount || 0), 0);
    
    // Group items by AF No first to detect ranges
    const itemsByAf: Record<string, typeof reportData> = {};
    // Also aggregate accounting entries
    const accountEntries: Record<string, { code: string, name: string, amount: number }> = {};

    reportData.forEach(item => {
      const af = item.afNo || 'Unspecified';
      if (!itemsByAf[af]) itemsByAf[af] = [];
      itemsByAf[af].push(item);

      // Accounting Entries
      const code = item.accountCode || 'No Code';
      const name = item.mainCategory || item.subCategory || 'Unspecified';
      const amt = item.amount || 0;
      
      // Create a unique key based on both code and name to ensure all unique Main Categories are listed
      const key = `${code}-${name}`;
      
      if (!accountEntries[key]) {
        accountEntries[key] = { code, name, amount: 0 };
      }
      accountEntries[key].amount += amt;
    });

    // Process ranges for AF List
    const afList: { 
        name: string, 
        minOr: string, 
        maxOr: string, 
        amount: number, 
        qty: number,
        begQty?: number,
        begFrom?: string,
        begTo?: string,
        issQty?: number,
        issFrom?: string,
        issTo?: string,
        endQty?: number,
        endFrom?: string,
        endTo?: string
    }[] = [];

    if ((startOr1 && endOr1) || (startOr2 && endOr2)) {
         const padLen = startOr1 ? startOr1.length : 7;

         // Helper to process range data
         const getRangeData = (s: string, e: string) => {
             const start = parseInt(s, 10);
             const end = parseInt(e, 10);
             const items = reportData.filter(item => {
                 if (!item.orNo) return false;
                 const itemOr = parseInt(item.orNo, 10);
                 if (!isNaN(itemOr) && !isNaN(start) && !isNaN(end)) {
                     return itemOr >= start && itemOr <= end;
                 }
                 return item.orNo >= s && item.orNo <= e;
             });
             const total = items.reduce((sum, item) => sum + (item.amount || 0), 0);
             return { total, count: items.length };
         };

          // Row 1
          if (startOr1 && endOr1) {
              const booklets = getBookletAccountability(startOr1, endOr1, padLen);
              booklets.forEach(entry => {
                  const entryData = getRangeData(entry.issFrom, entry.issTo);
                  afList.push({
                      name: 'A.F. NO. 51',
                      minOr: entry.issFrom,
                      maxOr: entry.issTo,
                      amount: entryData.total,
                      qty: entryData.count,
                      
                      begQty: entry.begQty,
                      begFrom: entry.begFrom,
                      begTo: entry.begTo,
                      issQty: entry.issQty,
                      issFrom: entry.issFrom,
                      issTo: entry.issTo,
                      endQty: entry.endQty,
                      endFrom: entry.endFrom,
                      endTo: entry.endTo
                  });
              });
          }

          // Row 2
          if (startOr2 && endOr2) {
              const booklets = getBookletAccountability(startOr2, endOr2, padLen);
              booklets.forEach(entry => {
                  const entryData = getRangeData(entry.issFrom, entry.issTo);
                  afList.push({
                      name: 'A.F. NO. 51',
                      minOr: entry.issFrom,
                      maxOr: entry.issTo,
                      amount: entryData.total,
                      qty: entryData.count,
                      
                      begQty: entry.begQty,
                      begFrom: entry.begFrom,
                      begTo: entry.begTo,
                      issQty: entry.issQty,
                      issFrom: entry.issFrom,
                      issTo: entry.issTo,
                      endQty: entry.endQty,
                      endFrom: entry.endFrom,
                      endTo: entry.endTo
                  });
              });
          }
     } else {
         Object.entries(itemsByAf).forEach(([af, items]) => {
             // Sort items by OR
             items.sort((a, b) => {
                  const orA = a.orNo || '';
                  const orB = b.orNo || '';
                  const numA = parseInt(orA, 10);
                  const numB = parseInt(orB, 10);
                  if (!isNaN(numA) && !isNaN(numB)) return numA - numB;
                  return orA.localeCompare(orB);
             });
     
             // Group into ranges
             let currentRange: { min: string, max: string, amount: number, qty: number, lastOrNum: number } | null = null;
             const rangesToProcess: { min: string, max: string }[] = [];
             
             for (const item of items) {
                 const orVal = item.orNo || '';
                 const orNum = parseInt(orVal, 10);
     
                 if (!currentRange) {
                     currentRange = { min: orVal, max: orVal, amount: 0, qty: 1, lastOrNum: isNaN(orNum) ? -999999 : orNum };
                 } else {
                     const isConsecutive = !isNaN(orNum) && !isNaN(currentRange.lastOrNum) && (orNum === currentRange.lastOrNum + 1);
                     
                     if (isConsecutive) {
                         currentRange.max = orVal;
                         currentRange.lastOrNum = orNum;
                     } else {
                         rangesToProcess.push({ min: currentRange.min, max: currentRange.max });
                         currentRange = { min: orVal, max: orVal, amount: 0, qty: 1, lastOrNum: isNaN(orNum) ? -999999 : orNum };
                     }
                 }
             }
             
             if (currentRange) {
                 rangesToProcess.push({ min: currentRange.min, max: currentRange.max });
             }

             // Process each detected range using getBookletAccountability
             rangesToProcess.forEach(range => {
                 const padLen = range.min.length || 7;
                 const booklets = getBookletAccountability(range.min, range.max, padLen);
                 booklets.forEach(entry => {
                     // Filter items for this booklet portion
                     const startVal = parseInt(entry.issFrom, 10);
                     const endVal = parseInt(entry.issTo, 10);
                     const entryItems = items.filter(item => {
                         const itemOr = parseInt(item.orNo || '', 10);
                         return !isNaN(itemOr) && itemOr >= startVal && itemOr <= endVal;
                     });
                     const entryAmount = entryItems.reduce((sum, item) => sum + (item.amount || 0), 0);

                     afList.push({
                         name: af,
                         minOr: entry.issFrom,
                         maxOr: entry.issTo,
                         amount: entryAmount,
                         qty: entry.issQty,
                         
                         begQty: entry.begQty,
                         begFrom: entry.begFrom,
                         begTo: entry.begTo,
                         issQty: entry.issQty,
                         issFrom: entry.issFrom,
                         issTo: entry.issTo,
                         endQty: entry.endQty,
                         endFrom: entry.endFrom,
                         endTo: entry.endTo
                     });
                 });
             });
         });
         
         // Sort final list by Name then by minOr
         afList.sort((a, b) => {
             const nameComp = a.name.localeCompare(b.name);
             if (nameComp !== 0) return nameComp;
             return parseInt(a.minOr, 10) - parseInt(b.minOr, 10);
         });
     }

    const sortedEntries = Object.values(accountEntries)
        .filter(entry => entry.amount !== 0)
        .sort((a, b) => a.code.localeCompare(b.code, undefined, { numeric: true }));

    // Date Logic
    const dateObj = reportData.length > 0 && reportData[0].date ? new Date(reportData[0].date) : new Date();
    const dateStr = dateObj.toLocaleDateString('en-PH', { month: 'long', year: 'numeric' });
    const certificationDateStr = new Date().toLocaleDateString('en-PH', { month: '2-digit', day: '2-digit', year: 'numeric' });
      
    // Report No Logic (YY-MM-??)
    const yy = dateObj.getFullYear().toString().slice(-2);
    const mm = (dateObj.getMonth() + 1).toString().padStart(2, '0');
    const reportNo = `${yy}-${mm}-`; 

    // Number to Words
    const numberToWords = (num: number): string => {
        const a = ['','One ','Two ','Three ','Four ','Five ','Six ','Seven ','Eight ','Nine ','Ten ','Eleven ','Twelve ','Thirteen ','Fourteen ','Fifteen ','Sixteen ','Seventeen ','Eighteen ','Nineteen '];
        const b = ['', '', 'Twenty','Thirty','Forty','Fifty','Sixty','Seventy','Eighty','Ninety'];

        const convert = (n: number): string => {
            if (n === 0) return '';
            if (n < 20) return a[n];
            if (n < 100) return b[Math.floor(n / 10)] + (n % 10 ? '-' + a[n % 10] : ' ');
            if (n < 1000) return a[Math.floor(n / 100)] + 'Hundred ' + (n % 100 ? convert(n % 100) : '');
            if (n < 1000000) return convert(Math.floor(n / 1000)) + 'Thousand ' + (n % 1000 ? convert(n % 1000) : '');
            return '';
        };

        const intPart = Math.floor(num);
        const centPart = Math.round((num - intPart) * 100);
        
        let words = intPart === 0 ? 'Zero ' : convert(intPart);
        words += 'Pesos';
        
        if (centPart > 0) {
            words += ` and ${convert(centPart)}Centavos`;
        } else {
            words += ' and Zero Centavos';
        }
        
        return `${words.trim()} only`;
    };

    const amountInWords = numberToWords(totalAmount);

    // Identify Signatories
    const collector = signatories.find(s => 
        s.position.toLowerCase().includes('rcc') || 
        s.position.toLowerCase().includes('collector') ||
        s.position.toLowerCase().includes('officer')
    ) || signatories[0] || { fullName: '______________________', position: 'Collector' };

    const treasurer = signatories.find(s => 
        s.position.toLowerCase().includes('treasurer')
    ) || signatories.find(s => s.id !== collector.id) || { fullName: '______________________', position: 'Treasurer' };

    const accountingHead = signatories.find(s => 
        s.position.toLowerCase().includes('accounting')
    ) || signatories[1] || { fullName: '______________________', position: 'Accounting Head' };

    const preparedBy = signatories.find(s =>
        s.id === 3
    ) || collector;

    const printWindow = window.open('', '_blank');
    if (printWindow) {
      printWindow.document.write(`
        <!DOCTYPE html> 
        <html lang="en"> 
        <head> 
        <meta charset="UTF-8"> 
        <title>Report of Collections and Deposits</title> 
        <style> 
            @page { 
                size: 8.5in 13in; 
                margin: 0.5in; 
            }
            body { 
                font-family: "Times New Roman", serif; 
                font-size: 11px; 
                margin: 0;
                padding: 0;
            } 
            .container { 
                width: 100%;
                margin: auto; 
                border: 2px solid #000; 
                padding: 5px; 
                box-sizing: border-box;
                height: 11.7in; /* Adjusted for external header */
                position: relative;
                display: flex;
                flex-direction: column;
            } 
            .header-text {
                text-align: center;
            }
            .header-text h3, .header-text h4 {
                margin: 2px 0;
                font-weight: bold;
            }
            .meta-table {
                width: 100%;
                border-collapse: collapse;
                margin-top: 5px;
                margin-bottom: 5px;
                border: none;
            }
            .meta-table td {
                padding: 2px 5px;
                border: none;
            }
            .main-table { 
                width: 100%; 
                border-collapse: collapse; 
                margin-top: 0; 
                border: 1px solid #000;
            } 
            .main-table th, .main-table td { 
                border: 1px solid #000; 
                padding: 3px; 
                font-size: 11px;
            } 
            .main-table th { 
                text-align: center; 
                background-color: #f0f0f0; /* Optional: light gray header */
            } 
            .section-header {
                font-weight: bold;
                margin-top: 5px;
                margin-bottom: 2px;
                text-transform: uppercase;
                font-size: 11px;
            }
            .right { text-align: right; } 
            .center { text-align: center; }
            .left { text-align: left; }
            .bold { font-weight: bold; }
            
            .page-break { 
                page-break-before: always; 
            }
            
            /* Page 2 Specifics */
            .page-2 {
                font-family: Arial, sans-serif;
                font-size: 10px;
            }
            .page-2 .main-table th {
                font-size: 10px;
            }
            .page-2 .main-table td {
                font-size: 10px;
            }

            /* Spacer to push content to bottom */
            .spacer {
                flex-grow: 1;
            }
            
            /* Signature area */
            .signature-table {
                width: 100%;
                margin-top: 20px;
                border: none;
            }
            .signature-table td {
                border: none;
                vertical-align: top;
                padding: 10px;
            }
            
            /* Divider lines */
            .line {
                border-bottom: 1px solid black;
                width: 80%;
                margin-top: 20px;
                margin-bottom: 5px;
            }
        </style> 
        </head> 
        <body> 
        
        <!-- PAGE 1 -->
        <div style="text-align: right; font-size: 12px; font-weight: bold; margin-bottom: 2px;">Appendix 34</div>
        <div class="container"> 
            <div style="position: absolute; top: 5px; right: 5px; font-size: 10px;">Page 1 of 2</div> 
        
            <div class="header-text" style="margin-top: 20px;">
                <h3>REPORT OF COLLECTIONS AND DEPOSITS</h3> 
                <h4>Concepcion, Romblon</h4> 
                <h4>LGU</h4> 
            </div>
        
            <table class="meta-table" style="border-top: 2px solid black; border-bottom: 2px solid black; margin-top: 10px;"> 
                <tr> 
                    <td width="60%">Fund: <b>GENERAL FUND</b></td> 
                    <td width="40%">Report No.: <b>${reportNo}</b></td> 
                </tr> 
                <tr> 
                    <td>Name of Accountable Officer: <b>${collector.fullName}</b></td> 
                    <td>Sheet No.: <b>01</b></td> 
                </tr> 
                <tr> 
                    <td></td> 
                    <td>Date: <b>${dateStr}</b></td> 
                </tr> 
            </table> 
        
            <div class="section-header">A. COLLECTIONS</div>
            <div class="section-header" style="text-indent: 15px;">1. For Collectors</div> 
        
            <table class="main-table"> 
                <thead>
                    <tr> 
                        <th rowspan="2" width="40%">Type (Form No.)</th> 
                        <th colspan="2">Official Receipt / Serial No.</th> 
                        <th rowspan="2" width="20%">Amount</th> 
                    </tr> 
                    <tr> 
                        <th width="20%">From</th> 
                        <th width="20%">To</th> 
                    </tr> 
                </thead>
                <tbody>
                    ${afList.map(item => `
                    <tr> 
                        <td>A.F. NO. 51</td> 
                        <td class="center">${item.minOr}</td> 
                        <td class="center">${item.maxOr}</td> 
                        <td class="right">${item.amount.toLocaleString('en-PH', { minimumFractionDigits: 2 })}</td> 
                    </tr> 
                    `).join('')}
                    
                    <!-- Filler rows for Section A.1 -->
                    ${Array(Math.max(0, 11 - afList.length)).fill(0).map(() => `
                    <tr><td>&nbsp;</td><td></td><td></td><td></td></tr> 
                    `).join('')}
                    
                    <tr> 
                        <td colspan="3" class="right bold">Total</td> 
                        <td class="right bold">${totalAmount.toLocaleString('en-PH', { minimumFractionDigits: 2 })}</td> 
                    </tr> 
                </tbody>
            </table> 
        
            <div class="section-header" style="text-indent: 15px; margin-top: 10px;">2. For Liquidating Officers / Treasurers</div> 
        
            <table class="main-table"> 
                <thead>
                    <tr> 
                        <th width="50%">Name of Accountable Officer</th> 
                        <th width="25%">Report No.</th> 
                        <th width="25%">Amount</th> 
                    </tr> 
                </thead>
                <tbody>
                    <tr> 
                        <td>${collector.fullName}</td> 
                        <td class="center">${reportNo}</td> 
                        <td class="right">${totalAmount.toLocaleString('en-PH', { minimumFractionDigits: 2 })}</td> 
                    </tr> 
                    <!-- Filler rows -->
                    ${Array(17).fill(0).map(() => `
                    <tr><td>&nbsp;</td><td></td><td></td></tr> 
                    `).join('')}
                    <tr> 
                        <td colspan="2" class="right bold">Total</td> 
                        <td class="right bold">${totalAmount.toLocaleString('en-PH', { minimumFractionDigits: 2 })}</td> 
                    </tr>
                </tbody>
            </table> 
        
            <div class="section-header" style="margin-top: 10px;">B. REMITTANCES/DEPOSITS</div>
            
            <table class="main-table"> 
                <thead>
                    <tr> 
                        <th width="50%">Accountable Officer / Bank</th> 
                        <th width="25%">Reference</th> 
                        <th width="25%">Amount</th> 
                    </tr> 
                </thead>
                <tbody>
                    <!-- Filler rows -->
                    ${Array(10).fill(0).map(() => `
                    <tr><td>&nbsp;</td><td></td><td></td></tr> 
                    `).join('')}
                    
                    <tr> 
                        <td colspan="2" class="right bold">TOTAL</td> 
                        <td class="center bold">-</td> 
                    </tr> 
                </tbody>
            </table> 
            
            <div class="spacer"></div>
        </div> 
        
        <div class="page-break"></div>

        <!-- PAGE 2 -->
        <div style="text-align: right; font-size: 12px; font-weight: bold; margin-bottom: 2px;">Appendix 34</div>
        <div class="container page-2">
            <div style="position: absolute; top: 5px; right: 5px; font-size: 10px;">Page 2 of 2</div> 
            
            <div class="section-header" style="margin-top: 20px;">C. ACCOUNTABILITY FOR ACCOUNTABLE FORMS</div> 
            
            <table class="main-table"> 
                <thead>
                    <tr> 
                        <th rowspan="2" width="20%">Name of Form & No.</th> 
                        <th colspan="3">Beginning Balance</th> 
                        <th colspan="3">Receipt</th> 
                        <th colspan="3">Issued</th> 
                        <th colspan="3">Ending Balance</th> 
                    </tr> 
                    <tr> 
                        <th width="5%">Qty</th><th width="7.5%">From</th><th width="7.5%">To</th> 
                        <th width="5%">Qty</th><th width="7.5%">From</th><th width="7.5%">To</th> 
                        <th width="5%">Qty</th><th width="7.5%">From</th><th width="7.5%">To</th> 
                        <th width="5%">Qty</th><th width="7.5%">From</th><th width="7.5%">To</th> 
                    </tr> 
                </thead>
                <tbody>
                    ${afList.map(item => `
                    <tr> 
                        <td>${item.name || 'A.F. NO. 51'}</td> 
                        <td>${item.begQty || ''}</td><td>${item.begFrom || ''}</td><td>${item.begTo || ''}</td> 
                        <td></td><td></td><td></td> 
                        <td class="center">${item.issQty || ''}</td><td class="center">${item.issFrom || ''}</td><td class="center">${item.issTo || ''}</td> 
                        <td>${item.endQty || ''}</td><td>${item.endFrom || ''}</td><td>${item.endTo || ''}</td> 
                    </tr> 
                    `).join('')}
                    ${afList.length === 0 ? `
                    <tr> 
                        <td>A.F. NO. 51</td> 
                        <td>50</td><td>0008451</td><td>0008500</td> 
                        <td></td><td></td><td></td> 
                        <td>3</td><td>0008451</td><td>0008453</td> 
                        <td>47</td><td>0008454</td><td>0008500</td> 
                    </tr>` : ''}
                    <!-- Filler rows -->
                    ${Array(Math.max(0, 5 - afList.length)).fill(0).map(() => `
                    <tr><td>&nbsp;</td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td></tr> 
                    `).join('')}
                </tbody>
            </table> 
            
            <div style="border: 2px solid black; padding: 5px; margin-top: 10px;">
                <div style="font-weight: bold; margin-bottom: 10px;">D. SUMMARY OF COLLECTIONS</div>
                <div style="display: flex; gap: 20px;">
                    <!-- Left Column: Summary -->
                    <div style="flex: 1;">
                        <table style="width: 100%; border-collapse: collapse; font-size: 11px;">
                            <tr>
                                <td colspan="3" style="padding-bottom: 5px;">Beginning Balance: ${dateStr}</td>
                            </tr>
                            <tr>
                                <td colspan="3">Add: Collections</td>
                            </tr>
                            <tr>
                                <td style="padding-left: 20px;">Cash</td>
                                <td style="text-align: right; width: 20px;">₱</td>
                                <td style="text-align: right; width: 100px;">${totalAmount.toLocaleString('en-PH', { minimumFractionDigits: 2 })}</td>
                            </tr>
                            <tr>
                                <td style="padding-left: 20px;">Check/s</td>
                                <td style="text-align: right;"></td>
                                <td style="text-align: right; border-bottom: 1px solid black;">-</td>
                            </tr>
                            <tr>
                                <td>Total</td>
                                <td style="text-align: right;">₱</td>
                                <td style="text-align: right; border-bottom: 1px solid black;">${totalAmount.toLocaleString('en-PH', { minimumFractionDigits: 2 })}</td>
                            </tr>
                            <tr>
                                <td>Less: Remittance/Deposit to Treasurer</td>
                                <td style="text-align: right;">₱</td>
                                <td style="text-align: right; border-bottom: 1px solid black;">${totalAmount.toLocaleString('en-PH', { minimumFractionDigits: 2 })}</td>
                            </tr>
                            <tr>
                                <td style="padding-top: 5px;">Balance</td>
                                <td style="text-align: right; padding-top: 5px;">₱</td>
                                <td style="text-align: right; border-bottom: 3px double black; padding-top: 5px;">-</td>
                            </tr>
                        </table>
                    </div>

                    <!-- Right Column: List of Checks -->
                    <div style="flex: 1;">
                         <div style="margin-bottom: 5px;">List of Checks :</div>
                         <table style="width: 100%; border-collapse: collapse; border: 2px solid black;">
                            <thead>
                                <tr>
                                    <th style="border: 1px solid black; width: 40%; height: 20px;"></th>
                                    <th style="border: 1px solid black; width: 30%;">Payee</th>
                                    <th style="border: 1px solid black; width: 30%;">Amount</th>
                                </tr>
                            </thead>
                            <tbody>
                                <tr><td style="border: 1px solid black; height: 20px;">&nbsp;</td><td style="border: 1px solid black;"></td><td style="border: 1px solid black;"></td></tr>
                                <tr><td style="border: 1px solid black; height: 20px;">&nbsp;</td><td style="border: 1px solid black;"></td><td style="border: 1px solid black;"></td></tr>
                                <tr><td style="border: 1px solid black; height: 20px;">&nbsp;</td><td style="border: 1px solid black;"></td><td style="border: 1px solid black;"></td></tr>
                            </tbody>
                         </table>
                         <div style="font-size: 10px; margin-top: 5px; text-align: center;">NOTE: Use additional sheet if necessary.</div>
                    </div>
                </div>

                <!-- Divider Line -->
                <div style="border-bottom: 2px solid black; margin: 20px 0;"></div>

                <!-- Certifications -->
                 <div style="display: flex; gap: 20px; font-size: 11px;">
                    <!-- Certification -->
                    <div style="flex: 1;">
                        <div style="font-weight: bold; margin-bottom: 10px;">CERTIFICATION:</div>
                        <p style="text-align: justify; margin-bottom: 40px; min-height: 40px;">
                            I hereby certify that the foregoing report of collections and deposits, and accountability for accountable forms is true and correct.
                        </p>
                        
                        <div style="display: flex; align-items: flex-end; gap: 10px;">
                            <div style="flex: 6; text-align: center;">
                                <div style="font-weight: bold; border-bottom: 1px solid black;">${collector.fullName}</div>
                                <div style="font-size: 10px;">${collector.position}</div>
                            </div>
                            <div style="flex: 4; text-align: center;">
                                <div style="border-bottom: 1px solid black;">${certificationDateStr}</div>
                                <div style="font-size: 10px;">Date</div>
                            </div>
                        </div>
                    </div>

                    <!-- Verification -->
                    <div style="flex: 1;">
                        <div style="font-weight: bold; margin-bottom: 10px;">VERIFICATION AND ACKNOWLEDGMENT:</div>
                        <p style="text-align: justify; margin-bottom: 25px; min-height: 40px; line-height: 1.6;">
                            I hereby certify that the foregoing report of collections has been verified and acknowledge receipt of (₱ ${totalAmount.toLocaleString('en-PH', { minimumFractionDigits: 2 })}) ${amountInWords}
                        </p>

                        <div style="display: flex; align-items: flex-end; gap: 10px;">
                             <div style="flex: 6; text-align: center;">
                                <div style="font-weight: bold; border-bottom: 1px solid black;">${treasurer.fullName}</div>
                                <div style="font-size: 10px;">${treasurer.position}</div>
                            </div>
                            <div style="flex: 4; text-align: center;">
                                <div style="border-bottom: 1px solid black;">&nbsp;</div>
                                <div style="font-size: 10px;">Date</div>
                            </div>
                        </div>
                    </div>
                </div>
            </div> 
            
            <div class="section-header" style="margin-top: 10px;">E. ACCOUNTING ENTRIES</div> 
            
            <table class="main-table"> 
                <thead>
                    <tr> 
                        <th width="40%">Particulars</th> 
                        <th width="20%">Account Code</th> 
                        <th width="20%">Debit</th> 
                        <th width="20%">Credit</th> 
                    </tr> 
                </thead>
                <tbody>
                    <tr> 
                        <td class="left">Cash in Local Treasury</td> 
                        <td class="center">1-01-01-010</td> 
                        <td class="right">₱ ${totalAmount.toLocaleString('en-PH', { minimumFractionDigits: 2 })}</td> 
                        <td class="right"></td> 
                    </tr> 
                    
                    ${sortedEntries.map(entry => `
                    <tr> 
                        <td class="left">${entry.name}</td> 
                        <td class="center">${entry.code}</td> 
                        <td class="right"></td> 
                        <td class="right">₱ ${entry.amount.toLocaleString('en-PH', { minimumFractionDigits: 2 })}</td> 
                    </tr> 
                    `).join('')}
                    
                    ${sortedEntries.length === 0 ? `
                    <tr><td class="left">Waterworks System Fees/Water-San Vicente</td><td class="center">4-02-02-090</td><td class="right"></td><td class="right">₱ 120.00</td></tr> 
                    <tr><td class="left">Waterworks System Fees/Water-Poblacion</td><td class="center">4-02-02-090</td><td class="right"></td><td class="right">₱ 200.00</td></tr> 
                    <tr><td class="left">Clearance and Certification Fees</td><td class="center">4-02-01-040</td><td class="right"></td><td class="right">₱ 120.00</td></tr> 
                    <tr><td class="left">Due to BIR</td><td class="center">2-02-01-010</td><td class="right"></td><td class="right">₱ 30.00</td></tr> 
                    ` : ''}

                    <!-- Filler rows -->
                    ${Array(Math.max(0, 24 - (1 + (sortedEntries.length || 4)))).fill(0).map(() => `
                    <tr><td>&nbsp;</td><td></td><td></td><td></td></tr> 
                    `).join('')}
                </tbody>
            </table> 
            
            <div class="spacer"></div>
            
            <table style="width: 100%; border: none; margin-bottom: 20px;">
                <tr>
                    <td width="50%" style="vertical-align: top;">
                        Prepared by:<br><br><br>
                        <strong>${preparedBy.fullName}</strong><br>
                        ${preparedBy.position}
                    </td>
                    <td width="50%" style="vertical-align: top;">
                        Certified Correct:<br><br><br>
                        <strong>${accountingHead.fullName}</strong><br>
                        ${accountingHead.position}
                    </td>
                </tr>
            </table>

        </div>
        </body> 
        </html>
      `);
      printWindow.document.close();
      printWindow.focus();
      printWindow.print();
    }
  };

  const handlePrintRptCover = () => {
    if (!rptFilterAf56Id) return;
    const afNo = rptFilterAf56Id;

    const filterByRange = (sOr: string | null, eOr: string | null) => {
      if (!sOr || !eOr) return [];
      return rptCollections.filter(item => {
        if (item.af56Id !== afNo || !item.orNumber) return false;
        
        const itemOr = parseInt(item.orNumber, 10);
        const start = parseInt(sOr, 10);
        const end = parseInt(eOr, 10);
        
        if (!isNaN(itemOr) && !isNaN(start) && !isNaN(end)) {
          return itemOr >= start && itemOr <= end;
        }
        return item.orNumber >= sOr && item.orNumber <= eOr;
      });
    };

    const range1Data = filterByRange(rptStartOr1, rptEndOr1);
    const range2Data = filterByRange(rptStartOr2, rptEndOr2);

    const calculateTotal = (data: RPTCollectionItem[]) => {
      return data.reduce((sum, item) => sum + (item.amount || 0), 0);
    };

    const r1Total = calculateTotal(range1Data);
    const r2Total = calculateTotal(range2Data);

    const printContent = (_rangeLabel: string, start: string | null, end: string | null, total: number) => {
      if (!start || !end) return '';

      // Calculate basic share (round up for odd cents)
      const basic = Math.round(total * 100 / 2) / 100;
      // SEF share is the remainder
      const sef = Number((total - basic).toFixed(2));

      return `
      <div class="column">
        <div class="header-box">
          ACCOUNTABLE FORM NO. 56
        </div>

        <div class="section">
          <div class="label">OR NO.</div>
          <div class="value">${start} - ${end}</div>
        </div>
        
        <div class="section">
          <div class="label">TOTAL AMOUNT:</div>
          <div class="value">${total.toLocaleString('en-PH', { minimumFractionDigits: 2 })}</div>
        </div>

        <div class="section breakdown">
          <div class="breakdown-item">
            <span class="red-label">BASIC:</span> <span class="amount-val">${basic.toLocaleString('en-PH', { minimumFractionDigits: 2 })}</span>
          </div>
          <div class="breakdown-item">
            <span class="red-label">SEF:</span> <span class="amount-val">${sef.toLocaleString('en-PH', { minimumFractionDigits: 2 })}</span>
          </div>
        </div>
      </div>
    `;
    };

    const printWindow = window.open('', '_blank');
    if (printWindow) {
      printWindow.document.write(`
        <html>
          <head>
            <title>Print RPT Cover - AF No. ${afNo}</title>
            <style>
              @page { size: Letter portrait; margin: 0.5in; }
              body { font-family: Arial, sans-serif; margin: 0; padding: 0; -webkit-print-color-adjust: exact; }
              .container { display: flex; width: 100%; height: 100vh; }
              .column { flex: 1; padding: 0 40px; box-sizing: border-box; display: flex; flex-direction: column; align-items: center; }
              .column:first-child { border-right: 1px dashed #ccc; }
              
              .header-box { 
                border: 3px double black; 
                padding: 5px 20px; 
                color: red; 
                font-weight: bold; 
                font-size: 18px;
                text-align: center; 
                margin-bottom: 50px;
                margin-top: 50px;
                width: 100%;
                box-sizing: border-box;
              }
              
              .section { margin-bottom: 40px; text-align: center; width: 100%; }
              
              .label { 
                font-size: 16px; 
                text-transform: uppercase; 
                margin-bottom: 15px; 
                color: black;
              }
              
              .value { 
                font-size: 28px; 
                font-weight: bold; 
                text-decoration: underline; 
              }

              .breakdown {
                font-size: 24px;
                display: flex;
                flex-direction: column;
                gap: 20px;
                margin-top: 20px;
              }

              .breakdown-item {
                display: flex;
                justify-content: center;
                gap: 10px;
              }

              .red-label {
                color: red;
                font-weight: bold;
              }

              .amount-val {
                color: black;
              }
            </style>
          </head>
          <body>
            <div class="container">
              ${printContent('Range 1', rptStartOr1, rptEndOr1, r1Total)}
              ${printContent('Range 2', rptStartOr2, rptEndOr2, r2Total)}
            </div>
          </body>
        </html>
      `);
      printWindow.document.close();
      printWindow.focus();
      printWindow.print();
    }
  };

  const handlePrintRptReport = (type: 'GENERAL' | 'SEF' | 'BOTH' = 'GENERAL') => {
    if (!rptFilterAf56Id) return;
    const afNo = rptFilterAf56Id;
    setRptReportMenuAnchor(null);

    const filterByRange = (sOr: string | null, eOr: string | null) => {
      if (!sOr || !eOr) return [];
      return rptCollections.filter(item => {
        if (item.af56Id !== afNo || !item.orNumber) return false;
        
        const itemOr = parseInt(item.orNumber, 10);
        const start = parseInt(sOr, 10);
        const end = parseInt(eOr, 10);
        
        if (!isNaN(itemOr) && !isNaN(start) && !isNaN(end)) {
          return itemOr >= start && itemOr <= end;
        }
        return item.orNumber >= sOr && item.orNumber <= eOr;
      });
    };

    const range1Data = filterByRange(rptStartOr1, rptEndOr1);
    const range2Data = filterByRange(rptStartOr2, rptEndOr2);
    
    // Merge data and sort by OR
    const allData = [...range1Data, ...range2Data].sort((a, b) => {
        const orA = parseInt(a.orNumber, 10);
        const orB = parseInt(b.orNumber, 10);
        return orA - orB;
    });

    if (allData.length === 0) return;

    // Helper to generate a single report HTML (General or SEF)
    const generateReportHtml = (fundType: 'GENERAL' | 'SEF', items: RPTCollectionItem[]) => {
      const isGeneral = fundType === 'GENERAL';
      const fundLabel = isGeneral ? 'GENERAL FUND' : 'SPECIAL EDUCATION FUND';

      // Keep Basic/SEF computation aligned with Print Cover:
      // split each selected range total (not each individual OR item).
      const splitByFund = (total: number) => {
        const basic = Math.round(total * 100 / 2) / 100;
        const sef = Number((total - basic).toFixed(2));
        return { basic, sef };
      };

      const getFundAmountForRange = (rangeItems: RPTCollectionItem[]) => {
        const rangeTotal = rangeItems.reduce((sum, item) => sum + (item.amount || 0), 0);
        const split = splitByFund(rangeTotal);
        return isGeneral ? split.basic : split.sef;
      };

      const r1FundTotal = getFundAmountForRange(range1Data);
      const r2FundTotal = getFundAmountForRange(range2Data);
      const totalAmount = Number((r1FundTotal + r2FundTotal).toFixed(2));

      // Date Logic
      let reportDateStr = new Date().toISOString().split('T')[0];
      if (items.length > 0) {
        const latestItem = items[items.length - 1];
        if (latestItem.date) {
          const s = String(latestItem.date).trim();
          
          // 1. Handle ISO strings with 'T'
          if (s.includes('T')) {
            const d = new Date(s);
            if (!isNaN(d.getTime())) {
              const month = String(d.getMonth() + 1).padStart(2, '0');
              const day = String(d.getDate()).padStart(2, '0');
              const year = d.getFullYear();
              reportDateStr = `${year}-${month}-${day}`;
            }
          } else {
            const cleanStr = s.split(' ')[0];
            // Try YYYY-MM-DD
            const ymd = cleanStr.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
            if (ymd) {
              const [, y, m, d] = ymd;
              reportDateStr = `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`;
            } else {
              // Try MM/DD/YYYY
              const mdy = cleanStr.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);
              if (mdy) {
                const [, m, d, y] = mdy;
                reportDateStr = `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`;
              }
            }
          }
        }
      }

      const [yFull, mNum] = reportDateStr.split('-');
      const monthNames = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
      const dateStr = `${monthNames[parseInt(mNum, 10) - 1]} ${yFull}`;
      
      const now = new Date();
      const certificationDateStr = `${String(now.getMonth() + 1).padStart(2, '0')}/${String(now.getDate()).padStart(2, '0')}/${now.getFullYear()}`;
      
      // Report No Logic
      const yy = yFull.slice(-2);
      const mm = mNum;
      const reportNo = `${yy}-${mm}-`; 

      // Number to Words
      const numberToWords = (num: number): string => {
        const a = ['','One ','Two ','Three ','Four ','Five ','Six ','Seven ','Eight ','Nine ','Ten ','Eleven ','Twelve ','Thirteen ','Fourteen ','Fifteen ','Sixteen ','Seventeen ','Eighteen ','Nineteen '];
        const b = ['', '', 'Twenty','Thirty','Forty','Fifty','Sixty','Seventy','Eighty','Ninety'];
        const convert = (n: number): string => {
            if (n === 0) return '';
            if (n < 20) return a[n];
            if (n < 100) return b[Math.floor(n / 10)] + (n % 10 ? '-' + a[n % 10] : ' ');
            if (n < 1000) return a[Math.floor(n / 100)] + 'Hundred ' + (n % 100 ? convert(n % 100) : '');
            if (n < 1000000) return convert(Math.floor(n / 1000)) + 'Thousand ' + (n % 1000 ? convert(n % 1000) : '');
            return '';
        };
        const intPart = Math.floor(num);
        const centPart = Math.round((num - intPart) * 100);
        let words = intPart === 0 ? 'Zero ' : convert(intPart);
        words += 'Pesos';
        if (centPart > 0) words += ` and ${convert(centPart)}Centavos`;
        else words += ' and Zero Centavos';
        return `${words.trim()} only`;
      };

      const amountInWords = numberToWords(totalAmount);

      // Identify Signatories
      const collector = signatories.find(s => 
          s.position.toLowerCase().includes('rcc') || 
          s.position.toLowerCase().includes('collector') ||
          s.position.toLowerCase().includes('officer')
      ) || signatories[0] || { fullName: '______________________', position: 'Collector' };

      const treasurer = signatories.find(s => 
          s.position.toLowerCase().includes('treasurer')
      ) || signatories.find(s => s.id !== collector.id) || { fullName: '______________________', position: 'Treasurer' };

      const accountingHead = signatories.find(s => 
          s.position.toLowerCase().includes('accounting')
      ) || signatories[1] || { fullName: '______________________', position: 'Accounting Head' };

      const preparedBy = signatories.find(s =>
          s.id === 3
      ) || collector;

      // Calculate OR Ranges for Section A
      // Simplified: Just use start/end ORs from the inputs since we know them
      const ranges = [];
      if (rptStartOr1 && rptEndOr1) {
        ranges.push({ min: rptStartOr1, max: rptEndOr1, amount: r1FundTotal });
      }
      if (rptStartOr2 && rptEndOr2) {
        ranges.push({ min: rptStartOr2, max: rptEndOr2, amount: r2FundTotal });
      }

      // Calculate Accountable Forms for Section C
      const afName = `A.F. NO. ${afNo}`;
      const afRows: string[] = [];

      const addAfRows = (sOr: string, eOr: string) => {
          const padLen = sOr.length || 7;
          const booklets = getBookletAccountability(sOr, eOr, padLen);
          booklets.forEach(entry => {
              afRows.push(`
                <tr> 
                    <td>${afName}</td> 
                    <td>${entry.begQty || ''}</td><td>${entry.begFrom || ''}</td><td>${entry.begTo || ''}</td> 
                    <td></td><td></td><td></td> 
                    <td class="center">${entry.issQty || ''}</td><td class="center">${entry.issFrom || ''}</td><td class="center">${entry.issTo || ''}</td> 
                    <td>${entry.endQty || ''}</td><td>${entry.endFrom || ''}</td><td>${entry.endTo || ''}</td> 
                </tr> 
              `);
          });
      };

      if (rptStartOr1 && rptEndOr1) addAfRows(rptStartOr1, rptEndOr1);
      if (rptStartOr2 && rptEndOr2) addAfRows(rptStartOr2, rptEndOr2);

      return `
        <div style="text-align: right; font-size: 12px; font-weight: bold; margin-bottom: 2px;">Appendix 34</div>
        <div class="container"> 
            <div style="position: absolute; top: 5px; right: 5px; font-size: 10px;">Page 1 of 2</div> 
        
            <div class="header-text" style="margin-top: 20px;">
                <h3>REPORT OF COLLECTIONS AND DEPOSITS</h3> 
                <h4>Concepcion, Romblon</h4> 
                <h4>LGU</h4> 
            </div>
        
            <table class="meta-table" style="border-top: 2px solid black; border-bottom: 2px solid black; margin-top: 10px;"> 
                <tr> 
                    <td width="60%">Fund: <b>${fundLabel}</b></td> 
                    <td width="40%">Report No.: <b>${reportNo}</b></td> 
                </tr> 
                <tr> 
                    <td>Name of Accountable Officer: <b>${collector.fullName}</b></td> 
                    <td>Sheet No.: <b>01</b></td> 
                </tr> 
                <tr> 
                    <td></td> 
                    <td>Date: <b>${dateStr}</b></td> 
                </tr> 
            </table> 
        
            <div class="section-header">A. COLLECTIONS</div>
            <div class="section-header" style="text-indent: 15px;">1. For Collectors</div> 
        
            <table class="main-table"> 
                <thead>
                    <tr> 
                        <th rowspan="2" width="40%">Type (Form No.)</th> 
                        <th colspan="2">Official Receipt / Serial No.</th> 
                        <th rowspan="2" width="20%">Amount</th> 
                    </tr> 
                    <tr> 
                        <th width="20%">From</th> 
                        <th width="20%">To</th> 
                    </tr> 
                </thead>
                <tbody>
                    ${ranges.map(r => `
                    <tr> 
                        <td>A.F. NO. 56</td> 
                        <td class="center">${r.min}</td> 
                        <td class="center">${r.max}</td> 
                        <td class="right">${r.amount.toLocaleString('en-PH', { minimumFractionDigits: 2 })}</td> 
                    </tr> 
                    `).join('')}
                    
                    ${Array(Math.max(0, 11 - ranges.length)).fill(0).map(() => `
                    <tr><td>&nbsp;</td><td></td><td></td><td></td></tr> 
                    `).join('')}
                    
                    <tr> 
                        <td colspan="3" class="right bold">Total</td> 
                        <td class="right bold">${totalAmount.toLocaleString('en-PH', { minimumFractionDigits: 2 })}</td> 
                    </tr> 
                </tbody>
            </table> 
        
            <div class="section-header" style="text-indent: 15px; margin-top: 10px;">2. For Liquidating Officers / Treasurers</div> 
        
            <table class="main-table"> 
                <thead>
                    <tr> 
                        <th width="50%">Name of Accountable Officer</th> 
                        <th width="25%">Report No.</th> 
                        <th width="25%">Amount</th> 
                    </tr> 
                </thead>
                <tbody>
                    <tr> 
                        <td>${collector.fullName}</td> 
                        <td class="center">${reportNo}</td> 
                        <td class="right">${totalAmount.toLocaleString('en-PH', { minimumFractionDigits: 2 })}</td> 
                    </tr> 
                    ${Array(17).fill(0).map(() => `
                    <tr><td>&nbsp;</td><td></td><td></td></tr> 
                    `).join('')}
                    <tr> 
                        <td colspan="2" class="right bold">Total</td> 
                        <td class="right bold">${totalAmount.toLocaleString('en-PH', { minimumFractionDigits: 2 })}</td> 
                    </tr>
                </tbody>
            </table> 
        
            <div class="section-header" style="margin-top: 10px;">B. REMITTANCES/DEPOSITS</div>
            
            <table class="main-table"> 
                <thead>
                    <tr> 
                        <th width="50%">Accountable Officer / Bank</th> 
                        <th width="25%">Reference</th> 
                        <th width="25%">Amount</th> 
                    </tr> 
                </thead>
                <tbody>
                    ${Array(10).fill(0).map(() => `
                    <tr><td>&nbsp;</td><td></td><td></td></tr> 
                    `).join('')}
                    <tr> 
                        <td colspan="2" class="right bold">TOTAL</td> 
                        <td class="center bold">-</td> 
                    </tr> 
                </tbody>
            </table> 
            <div class="spacer"></div>
        </div> 
        
        <div class="page-break"></div>

        <!-- PAGE 2 -->
        <div style="text-align: right; font-size: 12px; font-weight: bold; margin-bottom: 2px;">Appendix 34</div>
        <div class="container page-2">
            <div style="position: absolute; top: 5px; right: 5px; font-size: 10px;">Page 2 of 2</div> 
            
            <div class="section-header" style="margin-top: 20px;">C. ACCOUNTABILITY FOR ACCOUNTABLE FORMS</div> 
            
            <table class="main-table"> 
                <thead>
                    <tr> 
                        <th rowspan="2" width="20%">Name of Form & No.</th> 
                        <th colspan="3">Beginning Balance</th> 
                        <th colspan="3">Receipt</th> 
                        <th colspan="3">Issued</th> 
                        <th colspan="3">Ending Balance</th> 
                    </tr> 
                    <tr> 
                        <th width="5%">Qty</th><th width="7.5%">From</th><th width="7.5%">To</th> 
                        <th width="5%">Qty</th><th width="7.5%">From</th><th width="7.5%">To</th> 
                        <th width="5%">Qty</th><th width="7.5%">From</th><th width="7.5%">To</th> 
                        <th width="5%">Qty</th><th width="7.5%">From</th><th width="7.5%">To</th> 
                    </tr> 
                </thead>
                <tbody>
                    ${afRows.join('')}
                    ${Array(Math.max(0, 5 - afRows.length)).fill(0).map(() => `
                    <tr><td>&nbsp;</td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td></tr> 
                    `).join('')}
                </tbody>
            </table> 
            
            <div style="border: 2px solid black; padding: 5px; margin-top: 10px;">
                <div style="font-weight: bold; margin-bottom: 10px;">D. SUMMARY OF COLLECTIONS</div>
                <div style="display: flex; gap: 20px;">
                    <div style="flex: 1;">
                        <table style="width: 100%; border-collapse: collapse; font-size: 11px;">
                            <tr><td colspan="3" style="padding-bottom: 5px;">Beginning Balance: ${dateStr}</td></tr>
                            <tr><td colspan="3">Add: Collections</td></tr>
                            <tr>
                                <td style="padding-left: 20px;">Cash</td>
                                <td style="text-align: right; width: 20px;">₱</td>
                                <td style="text-align: right; width: 100px;">${totalAmount.toLocaleString('en-PH', { minimumFractionDigits: 2 })}</td>
                            </tr>
                            <tr>
                                <td style="padding-left: 20px;">Check/s</td>
                                <td style="text-align: right;"></td>
                                <td style="text-align: right; border-bottom: 1px solid black;">-</td>
                            </tr>
                            <tr>
                                <td>Total</td>
                                <td style="text-align: right;">₱</td>
                                <td style="text-align: right; border-bottom: 1px solid black;">${totalAmount.toLocaleString('en-PH', { minimumFractionDigits: 2 })}</td>
                            </tr>
                            <tr>
                                <td>Less: Remittance/Deposit to Treasurer</td>
                                <td style="text-align: right;">₱</td>
                                <td style="text-align: right; border-bottom: 1px solid black;">${totalAmount.toLocaleString('en-PH', { minimumFractionDigits: 2 })}</td>
                            </tr>
                            <tr>
                                <td style="padding-top: 5px;">Balance</td>
                                <td style="text-align: right; padding-top: 5px;">₱</td>
                                <td style="text-align: right; border-bottom: 3px double black; padding-top: 5px;">-</td>
                            </tr>
                        </table>
                    </div>
                    <div style="flex: 1;">
                         <div style="margin-bottom: 5px;">List of Checks :</div>
                         <table style="width: 100%; border-collapse: collapse; border: 2px solid black;">
                            <thead>
                                <tr>
                                    <th style="border: 1px solid black; width: 40%; height: 20px;"></th>
                                    <th style="border: 1px solid black; width: 30%;">Payee</th>
                                    <th style="border: 1px solid black; width: 30%;">Amount</th>
                                </tr>
                            </thead>
                            <tbody>
                                ${Array(3).fill(0).map(() => `<tr><td style="border: 1px solid black; height: 20px;">&nbsp;</td><td style="border: 1px solid black;"></td><td style="border: 1px solid black;"></td></tr>`).join('')}
                            </tbody>
                         </table>
                         <div style="font-size: 10px; margin-top: 5px; text-align: center;">NOTE: Use additional sheet if necessary.</div>
                    </div>
                </div>

                <div style="border-bottom: 2px solid black; margin: 20px 0;"></div>

                 <div style="display: flex; gap: 20px; font-size: 11px;">
                    <div style="flex: 1;">
                        <div style="font-weight: bold; margin-bottom: 10px;">CERTIFICATION:</div>
                        <p style="text-align: justify; margin-bottom: 40px; min-height: 40px;">
                            I hereby certify that the foregoing report of collections and deposits, and accountability for accountable forms is true and correct.
                        </p>
                        <div style="display: flex; align-items: flex-end; gap: 10px;">
                            <div style="flex: 6; text-align: center;">
                                <div style="font-weight: bold; border-bottom: 1px solid black;">${collector.fullName}</div>
                                <div style="font-size: 10px;">${collector.position}</div>
                            </div>
                            <div style="flex: 4; text-align: center;">
                                <div style="border-bottom: 1px solid black;">${certificationDateStr}</div>
                                <div style="font-size: 10px;">Date</div>
                            </div>
                        </div>
                    </div>

                    <div style="flex: 1;">
                        <div style="font-weight: bold; margin-bottom: 10px;">VERIFICATION AND ACKNOWLEDGMENT:</div>
                        <p style="text-align: justify; margin-bottom: 25px; min-height: 40px; line-height: 1.6;">
                            I hereby certify that the foregoing report of collections has been verified and acknowledge receipt of (₱ ${totalAmount.toLocaleString('en-PH', { minimumFractionDigits: 2 })}) ${amountInWords}
                        </p>
                        <div style="display: flex; align-items: flex-end; gap: 10px;">
                             <div style="flex: 6; text-align: center;">
                                <div style="font-weight: bold; border-bottom: 1px solid black;">${treasurer.fullName}</div>
                                <div style="font-size: 10px;">${treasurer.position}</div>
                            </div>
                            <div style="flex: 4; text-align: center;">
                                <div style="border-bottom: 1px solid black;">&nbsp;</div>
                                <div style="font-size: 10px;">Date</div>
                            </div>
                        </div>
                    </div>
                </div>
            </div> 
            
            <div class="section-header" style="margin-top: 10px;">E. ACCOUNTING ENTRIES</div> 
            
            <table class="main-table"> 
                <thead>
                    <tr> 
                        <th width="40%">Particulars</th> 
                        <th width="20%">Account Code</th> 
                        <th width="20%">Debit</th> 
                        <th width="20%">Credit</th> 
                    </tr> 
                </thead>
                <tbody>
                    <tr> 
                        <td class="left">Cash in Local Treasury</td> 
                        <td class="center">1-01-01-010</td> 
                        <td class="right">₱ ${totalAmount.toLocaleString('en-PH', { minimumFractionDigits: 2 })}</td> 
                        <td class="right"></td> 
                    </tr>
                    ${isGeneral ? `
                    <tr> 
                        <td class="left" style="padding-left: 20px;">Due to LGUs Barangay</td> 
                        <td class="center">2-02-01-070</td> 
                        <td class="right"></td> 
                        <td class="right"></td> 
                    </tr>
                    <tr> 
                        <td class="left" style="padding-left: 20px;">Due to LGUs Province</td> 
                        <td class="center">2-02-01-070</td> 
                        <td class="right"></td> 
                        <td class="right"></td> 
                    </tr>
                    <tr> 
                        <td class="left">Real Property Tax-Basic</td> 
                        <td class="center">4-01-02-040</td> 
                        <td class="right"></td> 
                        <td class="right"></td> 
                    </tr>
                    <tr> 
                        <td class="left">Discount on Real Property Tax- Basic</td> 
                        <td class="center">4-01-02-041</td> 
                        <td class="right"></td> 
                        <td class="right"></td> 
                    </tr>
                    <tr> 
                        <td class="left">Tax Revenue-Fines and Penalties Property Tax</td> 
                        <td class="center">4-01-05-020</td> 
                        <td class="right"></td> 
                        <td class="right"></td> 
                    </tr>
                    ` : `
                    `}
                    ${Array(isGeneral ? 16 : 21).fill(0).map(() => `<tr><td>&nbsp;</td><td></td><td></td><td></td></tr>`).join('')}
                </tbody>
            </table> 
            <div class="spacer"></div>
            
            <table style="width: 100%; border: none; margin-bottom: 20px;">
                <tr>
                    <td width="50%" style="vertical-align: top;">
                        Prepared by:<br><br><br>
                        <strong>${preparedBy.fullName}</strong><br>
                        ${preparedBy.position}
                    </td>
                    <td width="50%" style="vertical-align: top;">
                        Certified Correct:<br><br><br>
                        <strong>${accountingHead.fullName}</strong><br>
                        ${accountingHead.position}
                    </td>
                </tr>
            </table>
        </div>
      `;
    };

    const generalHtml = generateReportHtml('GENERAL', allData);
    const sefHtml = generateReportHtml('SEF', allData);

    let finalHtml = '';
    if (type === 'GENERAL') finalHtml = generalHtml;
    else if (type === 'SEF') finalHtml = sefHtml;
    else finalHtml = `${generalHtml}<div class="page-break"></div>${sefHtml}`;

    const printWindow = window.open('', '_blank');
    if (printWindow) {
      printWindow.document.write(`
        <!DOCTYPE html> 
        <html lang="en"> 
        <head> 
        <meta charset="UTF-8"> 
        <title>RPT Reports - ${type === 'BOTH' ? 'General & SEF' : type}</title> 
        <style> 
            @page { 
                size: 8.5in 13in; 
                margin: 0.5in; 
            }
            body { 
                font-family: "Times New Roman", serif; 
                font-size: 11px; 
                margin: 0;
                padding: 0;
            } 
            .container { 
                width: 100%;
                margin: auto; 
                border: 2px solid #000; 
                padding: 5px; 
                box-sizing: border-box;
                height: 11.7in; 
                position: relative;
                display: flex;
                flex-direction: column;
            } 
            .header-text { text-align: center; }
            .header-text h3, .header-text h4 { margin: 2px 0; font-weight: bold; }
            .meta-table { width: 100%; border-collapse: collapse; margin: 5px 0; border: none; }
            .meta-table td { padding: 2px 5px; border: none; }
            .main-table { width: 100%; border-collapse: collapse; margin-top: 0; border: 1px solid #000; } 
            .main-table th, .main-table td { border: 1px solid #000; padding: 3px; font-size: 11px; } 
            .main-table th { text-align: center; background-color: #f0f0f0; } 
            .section-header { font-weight: bold; margin: 5px 0 2px; text-transform: uppercase; font-size: 11px; }
            .right { text-align: right; } 
            .center { text-align: center; }
            .left { text-align: left; }
            .bold { font-weight: bold; }
            .page-break { page-break-before: always; }
            .spacer { flex-grow: 1; }
        </style> 
        </head> 
        <body> 
            ${finalHtml}
        </body> 
        </html>
      `);
      printWindow.document.close();
      printWindow.focus();
      printWindow.print();
    }
  };

  const handleExportToExcel = () => {
    if (filteredCollections.length === 0) return;

    // Helper to format date for Excel
    const formatExcelDate = (dateStr: string | undefined): string => {
      if (!dateStr) return '';
      const s = String(dateStr).trim();
      
      // 1. Handle ISO strings with 'T'
      if (s.includes('T')) {
        const d = new Date(s);
        if (!isNaN(d.getTime())) {
          const month = String(d.getMonth() + 1).padStart(2, '0');
          const day = String(d.getDate()).padStart(2, '0');
          const year = d.getFullYear();
          return `${month}/${day}/${year}`;
        }
      }

      // 2. Handle simple strings (remove time part if any)
      const cleanStr = s.split(' ')[0];

      // Try YYYY-MM-DD
      const ymd = cleanStr.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
      if (ymd) {
        const [, y, m, d] = ymd;
        return `${m.padStart(2, '0')}/${d.padStart(2, '0')}/${y}`;
      }
      
      // Try MM/DD/YYYY
      const mdy = cleanStr.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);
      if (mdy) {
        const [, m, d, y] = mdy;
        return `${m.padStart(2, '0')}/${d.padStart(2, '0')}/${y}`;
      }

      return cleanStr;
    };

    // Prepare data for export
    const exportData = filteredCollections.map(item => ({
      'OR Number': item.orNo,
      'Date': formatExcelDate(item.date),
      'Sub Category': item.subCategory,
      'Main Category': item.mainCategory,
      'Account Code': item.accountCode,
      'Amount': item.amount
    }));

    // Create worksheet
    const worksheet = XLSX.utils.json_to_sheet(exportData);
    
    // Create workbook
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Collection Details");

    // Generate filename based on date range if available
    let filename = "Collection_Details.xlsx";
    if (startDate && endDate) {
      filename = `Collection_Details_${startDate}_to_${endDate}.xlsx`;
    } else if (startDate) {
      filename = `Collection_Details_from_${startDate}.xlsx`;
    } else if (endDate) {
      filename = `Collection_Details_until_${endDate}.xlsx`;
    }

    // Export file
    XLSX.writeFile(workbook, filename);
  };

  return (
    <Box>
      <Box sx={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', mb: 4 }}>
        <Typography variant="h4" fontWeight="bold">Reports</Typography>
        <Box sx={{ display: 'flex', gap: 2, alignItems: 'center' }}>
          {tabValue === 1 && selectedAfNos.length === 1 && (
            <>
              <Box sx={{ width: 150 }}>
                <Autocomplete
                  size="small"
                  options={validOrs}
                  value={startOr1}
                  onChange={(_, newValue) => setStartOr1(newValue)}
                  renderInput={(params) => <TextField {...params} label="Start OR 1" />}
                />
              </Box>
              <Box sx={{ width: 150 }}>
                <Autocomplete
                  size="small"
                  options={validOrs}
                  value={endOr1}
                  onChange={(_, newValue) => setEndOr1(newValue)}
                  renderInput={(params) => <TextField {...params} label="End OR 1" />}
                />
              </Box>
              <Box sx={{ width: 150 }}>
                <Autocomplete
                  size="small"
                  options={validOrs}
                  value={startOr2}
                  onChange={(_, newValue) => setStartOr2(newValue)}
                  renderInput={(params) => <TextField {...params} label="Start OR 2" />}
                />
              </Box>
              <Box sx={{ width: 150 }}>
                <Autocomplete
                  size="small"
                  options={validOrs}
                  value={endOr2}
                  onChange={(_, newValue) => setEndOr2(newValue)}
                  renderInput={(params) => <TextField {...params} label="End OR 2" />}
                />
              </Box>
              <Button 
                variant="contained" 
                disabled={!((startOr1 && endOr1) || (startOr2 && endOr2))}
                onClick={handlePrintCover}
              >
                Print Cover
              </Button>
            </>
          )}
          {tabValue !== 2 && (
            <>
              <Button variant="contained" color="secondary" onClick={handlePrintReport}>
                Print Report
              </Button>
              <Button variant="outlined" startIcon={<Download />}>
                Export All
              </Button>
            </>
          )}

          {tabValue === 2 && rptFilterAf56Id && (
            <>
              <Box sx={{ width: 150 }}>
                <Autocomplete
                  size="small"
                  options={validRptOrs}
                  value={rptStartOr1}
                  onChange={(_, newValue) => setRptStartOr1(newValue)}
                  renderInput={(params) => <TextField {...params} label="Start OR 1" />}
                />
              </Box>
              <Box sx={{ width: 150 }}>
                <Autocomplete
                  size="small"
                  options={validRptOrs}
                  value={rptEndOr1}
                  onChange={(_, newValue) => setRptEndOr1(newValue)}
                  renderInput={(params) => <TextField {...params} label="End OR 1" />}
                />
              </Box>
              <Box sx={{ width: 150 }}>
                <Autocomplete
                  size="small"
                  options={validRptOrs}
                  value={rptStartOr2}
                  onChange={(_, newValue) => setRptStartOr2(newValue)}
                  renderInput={(params) => <TextField {...params} label="Start OR 2" />}
                />
              </Box>
              <Box sx={{ width: 150 }}>
                <Autocomplete
                  size="small"
                  options={validRptOrs}
                  value={rptEndOr2}
                  onChange={(_, newValue) => setRptEndOr2(newValue)}
                  renderInput={(params) => <TextField {...params} label="End OR 2" />}
                />
              </Box>
              <Button
                variant="contained"
                onClick={handlePrintRptCover}
                disabled={!rptStartOr1 || !rptEndOr1}
              >
                Print Cover
              </Button>
              <Button
                variant="contained"
                color="secondary"
                onClick={(e) => setRptReportMenuAnchor(e.currentTarget)}
                disabled={!rptStartOr1 || !rptEndOr1}
              >
                Print Report
              </Button>
              <Menu
                anchorEl={rptReportMenuAnchor}
                open={Boolean(rptReportMenuAnchor)}
                onClose={() => setRptReportMenuAnchor(null)}
              >
                <MenuItem onClick={() => handlePrintRptReport('GENERAL')}>General Report (Basic)</MenuItem>
                <MenuItem onClick={() => handlePrintRptReport('SEF')}>SEF Report (SEF)</MenuItem>
                <Divider />
                <MenuItem onClick={() => handlePrintRptReport('BOTH')}>Print Both</MenuItem>
              </Menu>
            </>
          )}
        </Box>
      </Box>

      <Paper sx={{ width: '100%', overflow: 'hidden' }}>
        <Tabs value={tabValue} onChange={handleTabChange} sx={{ borderBottom: 1, borderColor: 'divider', px: 2, pt: 2 }}>
          <Tab label="RCD Summaries" />
          <Tab label="Collection Details" />
          <Tab label="RPT Collections" />
        </Tabs>

        {isLoading ? (
          <Box sx={{ p: 4, display: 'flex', justifyContent: 'center' }}>
            <CircularProgress />
          </Box>
        ) : (
          <>
            {tabValue === 0 && (
              <Box sx={{ p: 3 }}>
                <Grid container spacing={3} sx={{ mb: 4 }}>
                  {/* AF No. Summary */}
                  <Grid size={{ xs: 12, md: 6 }}>
                    <Paper elevation={0} sx={{ p: 2, border: 1, borderColor: 'divider' }}>
                      <Typography variant="h6" gutterBottom fontWeight="bold">
                        Collection by Accountable Form
                      </Typography>
                      <TableContainer>
                        <Table size="small">
                          <TableHead>
                            <TableRow>
                              <TableCell sx={{ fontWeight: 'bold' }}>AF No.</TableCell>
                              <TableCell align="right" sx={{ fontWeight: 'bold' }}>Amount</TableCell>
                            </TableRow>
                          </TableHead>
                          <TableBody>
                            {afNoSummary.map((item) => (
                              <TableRow key={item.afNo}>
                                <TableCell>{item.afNo}</TableCell>
                                <TableCell align="right">
                                  ₱ {item.amount.toLocaleString('en-PH', { minimumFractionDigits: 2 })}
                                </TableCell>
                              </TableRow>
                            ))}
                            <TableRow>
                              <TableCell sx={{ fontWeight: 'bold' }}>Total</TableCell>
                              <TableCell align="right" sx={{ fontWeight: 'bold', color: 'primary.main' }}>
                                ₱ {afNoSummary.reduce((sum, item) => sum + item.amount, 0).toLocaleString('en-PH', { minimumFractionDigits: 2 })}
                              </TableCell>
                            </TableRow>
                          </TableBody>
                        </Table>
                      </TableContainer>
                    </Paper>
                  </Grid>

                  {/* Monthly Summary */}
                  <Grid size={{ xs: 12, md: 6 }}>
                    <Paper elevation={0} sx={{ p: 2, border: 1, borderColor: 'divider' }}>
                      <Typography variant="h6" gutterBottom fontWeight="bold">
                        Collection by Month & Year
                      </Typography>
                      <TableContainer>
                        <Table size="small">
                          <TableHead>
                            <TableRow>
                              <TableCell sx={{ fontWeight: 'bold' }}>Month & Year</TableCell>
                              <TableCell align="right" sx={{ fontWeight: 'bold' }}>Amount</TableCell>
                            </TableRow>
                          </TableHead>
                          <TableBody>
                            {monthlySummary.map((item) => (
                              <TableRow key={item.monthYear}>
                                <TableCell>{item.monthYear}</TableCell>
                                <TableCell align="right">
                                  ₱ {item.amount.toLocaleString('en-PH', { minimumFractionDigits: 2 })}
                                </TableCell>
                              </TableRow>
                            ))}
                            <TableRow>
                              <TableCell sx={{ fontWeight: 'bold' }}>Total</TableCell>
                              <TableCell align="right" sx={{ fontWeight: 'bold', color: 'primary.main' }}>
                                ₱ {monthlySummary.reduce((sum, item) => sum + item.amount, 0).toLocaleString('en-PH', { minimumFractionDigits: 2 })}
                              </TableCell>
                            </TableRow>
                          </TableBody>
                        </Table>
                      </TableContainer>
                    </Paper>
                  </Grid>

                  {/* Sub Category Summary */}
                  <Grid size={{ xs: 12, md: 6 }}>
                    <Paper elevation={0} sx={{ p: 2, border: 1, borderColor: 'divider' }}>
                      <Typography variant="h6" gutterBottom fontWeight="bold">
                        Collection by Sub Category
                      </Typography>
                      <TableContainer>
                        <Table size="small">
                          <TableHead>
                            <TableRow>
                              <TableCell sx={{ fontWeight: 'bold' }}>Sub Category</TableCell>
                              <TableCell align="right" sx={{ fontWeight: 'bold' }}>Amount</TableCell>
                            </TableRow>
                          </TableHead>
                          <TableBody>
                            {subCategorySummary.map((item) => (
                              <TableRow key={item.subCategory}>
                                <TableCell>{item.subCategory}</TableCell>
                                <TableCell align="right">
                                  ₱ {item.amount.toLocaleString('en-PH', { minimumFractionDigits: 2 })}
                                </TableCell>
                              </TableRow>
                            ))}
                            <TableRow>
                              <TableCell sx={{ fontWeight: 'bold' }}>Total</TableCell>
                              <TableCell align="right" sx={{ fontWeight: 'bold', color: 'primary.main' }}>
                                ₱ {subCategorySummary.reduce((sum, item) => sum + item.amount, 0).toLocaleString('en-PH', { minimumFractionDigits: 2 })}
                              </TableCell>
                            </TableRow>
                          </TableBody>
                        </Table>
                      </TableContainer>
                    </Paper>
                  </Grid>

                  {/* Main Category Summary */}
                  <Grid size={{ xs: 12, md: 6 }}>
                    <Paper elevation={0} sx={{ p: 2, border: 1, borderColor: 'divider' }}>
                      <Typography variant="h6" gutterBottom fontWeight="bold">
                        Collection by Main Category
                      </Typography>
                      <TableContainer>
                        <Table size="small">
                          <TableHead>
                            <TableRow>
                              <TableCell sx={{ fontWeight: 'bold' }}>Main Category</TableCell>
                              <TableCell align="right" sx={{ fontWeight: 'bold' }}>Amount</TableCell>
                            </TableRow>
                          </TableHead>
                          <TableBody>
                            {mainCategorySummary.map((item) => (
                              <TableRow key={item.mainCategory}>
                                <TableCell>{item.mainCategory}</TableCell>
                                <TableCell align="right">
                                  ₱ {item.amount.toLocaleString('en-PH', { minimumFractionDigits: 2 })}
                                </TableCell>
                              </TableRow>
                            ))}
                            <TableRow>
                              <TableCell sx={{ fontWeight: 'bold' }}>Total</TableCell>
                              <TableCell align="right" sx={{ fontWeight: 'bold', color: 'primary.main' }}>
                                ₱ {mainCategorySummary.reduce((sum, item) => sum + item.amount, 0).toLocaleString('en-PH', { minimumFractionDigits: 2 })}
                              </TableCell>
                            </TableRow>
                          </TableBody>
                        </Table>
                      </TableContainer>
                    </Paper>
                  </Grid>
                </Grid>
              </Box>
            )}

            {tabValue === 1 && (
              <>
                <Box sx={{ p: 2, borderBottom: 1, borderColor: 'divider', bgcolor: 'grey.50' }}>
                  <Box sx={{ display: 'flex', flexWrap: 'wrap', gap: 2, alignItems: 'center' }}>
                    <Box sx={{ flexGrow: 1, minWidth: '150px' }}>
                      <Autocomplete
                        multiple
                        size="small"
                        options={uniqueAfNos}
                        value={selectedAfNos}
                        onChange={(_, newValue) => setSelectedAfNos(newValue)}
                        renderInput={(params) => <TextField {...params} label="AF No." placeholder="All AF No." />}
                      />
                    </Box>

                    <Box sx={{ flexGrow: 2, minWidth: '250px' }}>
                      <Autocomplete
                        multiple
                        size="small"
                        options={uniqueSubCategories}
                        value={selectedSubCategories}
                        onChange={(_, newValue) => setSelectedSubCategories(newValue)}
                        renderInput={(params) => <TextField {...params} label="Sub Category" placeholder="All Sub Categories" />}
                      />
                    </Box>
                    <Box sx={{ flexGrow: 2, minWidth: '250px' }}>
                      <Autocomplete
                        multiple
                        size="small"
                        options={uniqueMainCategories}
                        value={selectedMainCategories}
                        onChange={(_, newValue) => setSelectedMainCategories(newValue)}
                        renderInput={(params) => <TextField {...params} label="Main Category" placeholder="All Main Categories" />}
                      />
                    </Box>
                    <Box sx={{ flexGrow: 1, minWidth: '150px' }}>
                      <TextField
                        type="date"
                        label="Start Date"
                        size="small"
                        fullWidth
                        InputLabelProps={{ shrink: true }}
                        value={startDate}
                        onChange={(e) => setStartDate(e.target.value)}
                      />
                    </Box>
                    <Box sx={{ flexGrow: 1, minWidth: '150px' }}>
                      <TextField
                        type="date"
                        label="End Date"
                        size="small"
                        fullWidth
                        InputLabelProps={{ shrink: true }}
                        value={endDate}
                        onChange={(e) => setEndDate(e.target.value)}
                      />
                    </Box>

                    <Button 
                      variant="contained" 
                      color="success" 
                      startIcon={<FileDownload />}
                      onClick={handleExportToExcel}
                      disabled={filteredCollections.length === 0}
                      sx={{ height: 40 }}
                    >
                      Export to Excel
                    </Button>

                    <Box sx={{ flexGrow: 1, display: 'flex', justifyContent: 'flex-end', alignItems: 'center', minWidth: '300px', ml: 'auto' }}>
                      <Typography variant="h6" fontWeight="bold" color="text.secondary" sx={{ mr: 2 }}>
                        Total Amount:
                      </Typography>
                      <Typography variant="h5" fontWeight="bold" color="primary.main">
                        ₱ {totalFilteredAmount.toLocaleString('en-PH', { minimumFractionDigits: 2 })}
                      </Typography>
                    </Box>
                  </Box>
                </Box>
                <TableContainer sx={{ maxHeight: 700 }}>
                  <Table stickyHeader>
                    <TableHead>
                      <TableRow>
                        <TableCell sx={{ fontWeight: 'bold' }}>Sub Category</TableCell>
                        <TableCell sx={{ fontWeight: 'bold' }}>Main Category</TableCell>
                        <TableCell sx={{ fontWeight: 'bold' }}>Account Code</TableCell>
                        <TableCell align="right" sx={{ fontWeight: 'bold' }}>Amount</TableCell>
                      </TableRow>
                    </TableHead>
                    <TableBody>
                      {filteredCollections
                        .slice(page * rowsPerPage, page * rowsPerPage + rowsPerPage)
                        .map((item) => (
                        <TableRow key={item.id} hover>
                          <TableCell>{item.subCategory}</TableCell>
                          <TableCell>{item.mainCategory}</TableCell>
                          <TableCell sx={{ fontFamily: 'monospace', bgcolor: 'grey.50' }}>{item.accountCode}</TableCell>
                          <TableCell align="right" sx={{ fontWeight: 'bold' }}>
                            ₱ {(item.amount || 0).toLocaleString('en-PH', { minimumFractionDigits: 2 })}
                          </TableCell>
                        </TableRow>
                      ))}
                      {filteredCollections.length === 0 && (
                        <TableRow>
                          <TableCell colSpan={4} align="center" sx={{ py: 8 }}>
                            <Typography color="text.secondary">No collection entries found matching filters.</Typography>
                          </TableCell>
                        </TableRow>
                      )}
                    </TableBody>
                  </Table>
                </TableContainer>
                <TablePagination
                  rowsPerPageOptions={[10, 25, 50]}
                  component="div"
                  count={filteredCollections.length}
                  rowsPerPage={rowsPerPage}
                  page={page}
                  onPageChange={handleChangePage}
                  onRowsPerPageChange={handleChangeRowsPerPage}
                  labelRowsPerPage="Entries per page:"
                />
              </>
            )}

            {tabValue === 2 && (
              <>
                <Box sx={{ p: 2, borderBottom: 1, borderColor: 'divider', bgcolor: 'grey.50' }}>
                  <Box sx={{ display: 'flex', flexWrap: 'wrap', gap: 2, alignItems: 'center' }}>
                    <Box sx={{ flexGrow: 1, minWidth: '200px' }}>
                      <Autocomplete
                        options={uniqueRptAf56Ids}
                        value={rptFilterAf56Id}
                        onChange={(_, newValue) => setRptFilterAf56Id(newValue)}
                        renderInput={(params) => <TextField {...params} label="AF56 ID" size="small" />}
                      />
                    </Box>
                    <Box sx={{ flexGrow: 1, minWidth: '150px' }}>
                      <TextField
                        type="date"
                        label="Start Date"
                        size="small"
                        fullWidth
                        InputLabelProps={{ shrink: true }}
                        value={rptStartDate}
                        onChange={(e) => setRptStartDate(e.target.value)}
                      />
                    </Box>
                    <Box sx={{ flexGrow: 1, minWidth: '150px' }}>
                      <TextField
                        type="date"
                        label="End Date"
                        size="small"
                        fullWidth
                        InputLabelProps={{ shrink: true }}
                        value={rptEndDate}
                        onChange={(e) => setRptEndDate(e.target.value)}
                      />
                    </Box>
                    <Button 
                      variant="outlined" 
                      color="secondary" 
                      startIcon={<Clear />} 
                      onClick={() => {
                        setRptFilterAf56Id(null);
                        setRptStartDate('');
                        setRptEndDate('');
                      }}
                    >
                      Clear
                    </Button>

                    <Box sx={{ flexGrow: 1, display: 'flex', justifyContent: 'flex-end', alignItems: 'center', minWidth: '300px', ml: 'auto' }}>
                      <Typography variant="h6" fontWeight="bold" color="text.secondary" sx={{ mr: 2 }}>
                        Total Amount:
                      </Typography>
                      <Typography variant="h5" fontWeight="bold" color="primary.main">
                        ₱ {filteredRptCollections.reduce((sum, item) => sum + (item.amount || 0), 0).toLocaleString('en-PH', { minimumFractionDigits: 2 })}
                      </Typography>
                    </Box>
                  </Box>
                </Box>
                <TableContainer sx={{ maxHeight: 700 }}>
                  <Table stickyHeader>
                    <TableHead>
                      <TableRow>
                        <TableCell sx={{ fontWeight: 'bold' }}>Date</TableCell>
                        <TableCell sx={{ fontWeight: 'bold' }}>OR Number</TableCell>
                        <TableCell sx={{ fontWeight: 'bold' }}>Payor</TableCell>
                        <TableCell sx={{ fontWeight: 'bold' }}>Barangay</TableCell>
                        <TableCell align="right" sx={{ fontWeight: 'bold' }}>Amount</TableCell>
                        <TableCell sx={{ fontWeight: 'bold' }}>Remarks</TableCell>
                      </TableRow>
                    </TableHead>
                    <TableBody>
                      {filteredRptCollections
                        .slice(page * rowsPerPage, page * rowsPerPage + rowsPerPage)
                        .map((item) => (
                        <TableRow key={item.id} hover>
                          <TableCell>{item.date}</TableCell>
                          <TableCell>{item.orNumber}</TableCell>
                          <TableCell>{item.payor}</TableCell>
                          <TableCell>{item.barangay}</TableCell>
                          <TableCell align="right" sx={{ fontWeight: 'bold' }}>
                            ₱ {(item.amount || 0).toLocaleString('en-PH', { minimumFractionDigits: 2 })}
                          </TableCell>
                          <TableCell>{item.remarks || '-'}</TableCell>
                        </TableRow>
                      ))}
                      {filteredRptCollections.length === 0 && (
                        <TableRow>
                          <TableCell colSpan={6} align="center" sx={{ py: 8 }}>
                            <Typography color="text.secondary">No RPT collections found matching filters.</Typography>
                          </TableCell>
                        </TableRow>
                      )}
                    </TableBody>
                  </Table>
                </TableContainer>
                <TablePagination
                  rowsPerPageOptions={[10, 25, 50]}
                  component="div"
                  count={filteredRptCollections.length}
                  rowsPerPage={rowsPerPage}
                  page={page}
                  onPageChange={handleChangePage}
                  onRowsPerPageChange={handleChangeRowsPerPage}
                  labelRowsPerPage="Entries per page:"
                />
              </>
            )}
          </>
        )}
      </Paper>
    </Box>
  );
};
