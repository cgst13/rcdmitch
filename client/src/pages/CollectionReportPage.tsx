import React, { useEffect, useMemo, useState, useRef } from 'react';
import {
  Box,
  Container,
  Typography,
  Paper,
  Grid,
  TextField,
  Button,
  Stack,
  Autocomplete,
  Table,
  TableHead,
  TableBody,
  TableRow,
  TableCell,
  TableContainer,
  Card,
  CardHeader,
  CardContent,
  IconButton,
  InputAdornment,
  Tooltip,
  Backdrop,
  CircularProgress,
  TablePagination,
  TableSortLabel
} from '@mui/material';
import { 
  AddCircleOutline, 
  DeleteOutline, 
  ReceiptLong, 
  PersonOutline, 
  AttachMoney, 
  Event, 
  Notes, 
  Save,
  Edit,
  Search,
  Clear,
  ExpandLess,
  ExpandMore
} from '@mui/icons-material';
import { ConfirmDialog } from '../components/ConfirmDialog';
import { Notification } from '../components/Notification';
import { getAccountCodes, getCollectionEntries, saveCollectionEntryBulk, updateCollectionEntry, deleteCollectionEntry, type CollectionEntry } from '../services/googleSheets';
import type { AccountCode } from '../types/rcd';

const getNextOrNo = (currentOrNo: string): string => {
  let nextOrNo = currentOrNo.trim();
  const matchNum = nextOrNo.match(/(\d+)$/);
  if (matchNum) {
    const numStr = matchNum[1];
    const nextNum = parseInt(numStr, 10) + 1;
    const nextNumStr = String(nextNum).padStart(8, '0');
    nextOrNo = nextOrNo.substring(0, matchNum.index) + nextNumStr;
  } else if (nextOrNo && !isNaN(Number(nextOrNo))) {
    nextOrNo = String(Number(nextOrNo) + 1).padStart(8, '0');
  }
  return nextOrNo;
};

const parseDateToInputFormat = (dateStr: string): string => {
  const now = new Date();
  const today = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
  
  if (!dateStr) return today;
  
  const s = String(dateStr).trim();

  // 1. Handle ISO strings with 'T'
  if (s.includes('T')) {
    const d = new Date(s);
    if (!isNaN(d.getTime())) {
      const month = String(d.getMonth() + 1).padStart(2, '0');
      const day = String(d.getDate()).padStart(2, '0');
      const year = d.getFullYear();
      return `${year}-${month}-${day}`;
    }
  }

  // 2. Handle simple strings
  const cleanStr = s.split(' ')[0];

  // Try YYYY-MM-DD
  const ymd = cleanStr.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
  if (ymd) {
    const [, y, m, d] = ymd;
    return `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`;
  }
  
  // Try MM/DD/YYYY
  const mdy = cleanStr.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);
  if (mdy) {
    const [, m, d, y] = mdy;
    return `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`;
  }

  return today;
};

const formatDisplayDate = (dateStr: string): string => {
  if (!dateStr) return '';
  
  const s = String(dateStr).trim();

  // 1. Handle ISO strings with 'T'
  if (s.includes('T')) {
    const d = new Date(s);
    if (!isNaN(d.getTime())) {
      const month = String(d.getMonth() + 1).padStart(2, '0');
      const day = String(d.getDate()).padStart(2, '0');
      const year = d.getFullYear();
      return `${month}-${day}-${year}`;
    }
  }

  // 2. Handle simple strings
  const cleanStr = s.split(' ')[0];

  // Try YYYY-MM-DD
  const ymd = cleanStr.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
  if (ymd) {
    const [, y, m, d] = ymd;
    return `${m.padStart(2, '0')}-${d.padStart(2, '0')}-${y}`;
  }
  
  // Try MM/DD/YYYY
  const mdy = cleanStr.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);
  if (mdy) {
    const [, m, d, y] = mdy;
    return `${m.padStart(2, '0')}-${d.padStart(2, '0')}-${y}`;
  }

  return cleanStr;
};

export const CollectionReportPage: React.FC = () => {
  const payorRef = useRef<HTMLInputElement>(null);
  const subCategoryRefs = useRef<(HTMLInputElement | null)[]>([]);
  const [showEntryForm, setShowEntryForm] = useState(false);
  const [loading, setLoading] = useState(false);
  const [notification, setNotification] = useState<{
    open: boolean;
    message: string;
    severity: 'success' | 'error' | 'info' | 'warning';
  }>({ open: false, message: '', severity: 'info' });

  // Delete Dialog State
  const [deleteDialogOpen, setDeleteDialogOpen] = useState(false);
  const [itemToDelete, setItemToDelete] = useState<number | null>(null);

  const [editingId, setEditingId] = useState<number | null>(null);

  const [searchTerm, setSearchTerm] = useState('');
  const [filterAfNo, setFilterAfNo] = useState<string | null>(null);
  const [filterSubCategory, setFilterSubCategory] = useState<string | null>(null);
  const [filterDate, setFilterDate] = useState('');

  const [accountCodes, setAccountCodes] = useState<AccountCode[]>([]);
  const [items, setItems] = useState<CollectionEntry[]>([]);

  const [page, setPage] = useState(0);
  const [rowsPerPage, setRowsPerPage] = useState(10);
  const [sortBy, setSortBy] = useState<'orNo' | 'date'>('date');
  const [sortOrder, setSortOrder] = useState<'asc' | 'desc'>('desc');
  const [expandedOrNos, setExpandedOrNos] = useState<string[]>([]);

  const [form, setForm] = useState({
    afNo: '',
    orNo: '',
    payor: '',
    date: new Date().toISOString().split('T')[0],
    remarks: ''
  });
  const [charges, setCharges] = useState<Array<{
    subCategory: string;
    mainCategory: string;
    accountCode: string;
    amount: string;
  }>>([{ subCategory: '', mainCategory: '', accountCode: '', amount: '' }]);

  const handleCloseNotification = () => {
    setNotification(prev => ({ ...prev, open: false }));
  };

  const handleChangePage = (_event: unknown, newPage: number) => {
    setPage(newPage);
  };

  const handleChangeRowsPerPage = (event: React.ChangeEvent<HTMLInputElement>) => {
    setRowsPerPage(parseInt(event.target.value, 10));
    setPage(0);
  };

  useEffect(() => {
    const loadData = async () => {
      setLoading(true);
      try {
        const [codes, entries] = await Promise.all([
          getAccountCodes(),
          getCollectionEntries()
        ]);
        
        setAccountCodes(codes);
        // Sort entries by ID descending so newest are first
        setItems(entries.sort((a, b) => Number(b.id) - Number(a.id)));
        
        if (entries.length > 0) {
          // Find the entry with largest ID for next OR number calculation
          const lastEntry = entries[0]; // Since it's sorted desc, first is latest
          const nextOrNo = getNextOrNo(lastEntry.orNo);
          
          setForm(prev => ({
            ...prev,
            afNo: lastEntry.afNo,
            orNo: nextOrNo,
            date: parseDateToInputFormat(lastEntry.date),
            payor: '',
            remarks: ''
          }));
        }
      } catch (error) {
        console.error("Failed to load data", error);
        setNotification({
          open: true,
          message: 'Failed to load initial data. Please refresh.',
          severity: 'error'
        });
      } finally {
        setLoading(false);
      }
    };
    loadData();
  }, []);

  const mainCategories = useMemo(
    () => Array.from(new Set(accountCodes.map(c => c.mainCategory))).filter(Boolean).sort(),
    [accountCodes]
  );
  const subCategories = useMemo(
    () => Array.from(new Set(accountCodes.map(c => c.subCategory))).filter(Boolean).sort(),
    [accountCodes]
  );
  const accountCodeOptionsBySub = useMemo(() => {
    const map = new Map<string, string[]>();
    accountCodes.forEach(c => {
      const curr = map.get(c.subCategory) || [];
      map.set(c.subCategory, [...curr, c.code]);
    });
    return map;
  }, [accountCodes]);

  const uniquePayors = useMemo(() => {
    const payors = new Set(items.map(i => i.payor).filter(p => p && p.trim() !== ''));
    return Array.from(payors).sort();
  }, [items]);

  const uniqueAfNos = useMemo(() => {
    const afs = new Set(items.map(i => i.afNo).filter(Boolean));
    return Array.from(afs).sort();
  }, [items]);

  const filteredItems = useMemo(() => {
    const result = items.filter(item => {
      // Search Term (General)
      const searchLower = searchTerm.toLowerCase();
      const matchesSearch = 
        !searchTerm ||
        (item.orNo || '').toLowerCase().includes(searchLower) ||
        (item.payor || '').toLowerCase().includes(searchLower) ||
        (item.remarks || '').toLowerCase().includes(searchLower) ||
        (item.afNo || '').toLowerCase().includes(searchLower);

      // Filters
      const matchesAf = filterAfNo ? item.afNo === filterAfNo : true;
      const matchesSub = filterSubCategory ? item.subCategory === filterSubCategory : true;
      
      let matchesDate = true;
      if (filterDate) {
        // filterDate is YYYY-MM
        // item.date could be YYYY-MM-DD or M/D/YYYY or MM/DD/YYYY
        const normalizedDate = parseDateToInputFormat(item.date); // Returns YYYY-MM-DD
        matchesDate = normalizedDate.startsWith(filterDate);
      }

      return matchesSearch && matchesAf && matchesSub && matchesDate;
    });

    return result.sort((a, b) => {
      if (sortBy === 'orNo') {
        const orA = parseInt(a.orNo || '0', 10);
        const orB = parseInt(b.orNo || '0', 10);
        return sortOrder === 'asc' ? orA - orB : orB - orA;
      } else {
        const dateA = new Date(a.date).getTime();
        const dateB = new Date(b.date).getTime();
        return sortOrder === 'asc' ? dateA - dateB : dateB - dateA;
      }
    });
  }, [items, searchTerm, filterAfNo, filterSubCategory, filterDate, sortBy, sortOrder]);

  const groupedItems = useMemo(() => {
    const groups = new Map<string, {
      orNo: string;
      afNo: string;
      payor: string;
      date: string;
      remarks: string;
      charges: CollectionEntry[];
      total: number;
    }>();

    filteredItems.forEach(item => {
      const key = item.orNo || 'UNKNOWN';
      const existing = groups.get(key);
      if (existing) {
        existing.charges.push(item);
        existing.total += item.amount || 0;
      } else {
        groups.set(key, {
          orNo: key,
          afNo: item.afNo || '',
          payor: item.payor || '',
          date: item.date || '',
          remarks: item.remarks || '',
          charges: [item],
          total: item.amount || 0
        });
      }
    });

    return Array.from(groups.values());
  }, [filteredItems]);

  const visibleGroups = useMemo(
    () => groupedItems.slice(page * rowsPerPage, page * rowsPerPage + rowsPerPage),
    [groupedItems, page, rowsPerPage]
  );

  const toggleGroup = (orNo: string) => {
    setExpandedOrNos(prev => 
      prev.includes(orNo) ? prev.filter(item => item !== orNo) : [...prev, orNo]
    );
  };

  const handleEdit = (row: any) => {
    setEditingId(row.id);
    setForm({
      afNo: row.afNo || '',
      orNo: row.orNo || '',
      payor: row.payor || '',
      date: parseDateToInputFormat(row.date),
      remarks: row.remarks || ''
    });
    setCharges([{
      subCategory: row.subCategory || '',
      mainCategory: row.mainCategory || '',
      accountCode: row.accountCode || '',
      amount: String(row.amount || '')
    }]);
    setShowEntryForm(true);
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  const confirmDelete = (id: number) => {
    setItemToDelete(id);
    setDeleteDialogOpen(true);
  };

  const handleDelete = async () => {
    if (itemToDelete === null) return;
    
    setLoading(true);
    try {
      const success = await deleteCollectionEntry(itemToDelete);
      if (success) {
        setItems(prev => prev.filter(item => item.id !== itemToDelete));
        setNotification({
          open: true,
          message: 'Entry deleted successfully',
          severity: 'success'
        });
      } else {
        throw new Error('Failed to delete');
      }
    } catch (error) {
      console.error('Delete error:', error);
      setNotification({
        open: true,
        message: 'Failed to delete entry',
        severity: 'error'
      });
    } finally {
      setLoading(false);
      setItemToDelete(null);
    }
  };

  const handleCancelEdit = () => {
    setEditingId(null);
    setShowEntryForm(false);
    // Reset form to next OR logic
    if (items.length > 0) {
      // items is sorted desc by ID
      const lastEntry = items[0]; 
      const nextOrNo = getNextOrNo(lastEntry.orNo);
      setForm(prev => ({
        ...prev,
        afNo: lastEntry.afNo,
        orNo: nextOrNo,
        date: parseDateToInputFormat(lastEntry.date),
        payor: '',
        remarks: ''
      }));
    } else {
      setForm(prev => ({
        ...prev,
        orNo: '',
        date: new Date().toISOString().split('T')[0],
        payor: '',
        remarks: ''
      }));
    }
    setCharges([{ subCategory: '', mainCategory: '', accountCode: '', amount: '' }]);
  };

  const addItem = async () => {
    const header = {
      afNo: (form.afNo || '').trim(),
      orNo: (form.orNo || '').trim(),
      payor: (form.payor || '').trim(),
      date: form.date,
      remarks: (form.remarks || '').trim()
    };
    const preparedCharges = charges
      .map(c => {
        const sub = (c.subCategory || '').trim();
        const match = accountCodes.find(a => a.subCategory === sub);
        return {
          subCategory: sub,
          mainCategory: c.mainCategory || (match ? match.mainCategory : ''),
          accountCode: c.accountCode || (match ? match.code : ''),
          amount: Number(c.amount) || 0
        };
      })
      .filter(c => c.subCategory);
    
    if (preparedCharges.length === 0) {
      setNotification({
        open: true,
        message: 'Please add at least one charge with a Sub Category.',
        severity: 'warning'
      });
      return;
    }

    if (editingId) {
      const charge = preparedCharges[0];
      const entry: CollectionEntry = {
        id: editingId,
        afNo: header.afNo,
        orNo: header.orNo,
        payor: header.payor,
        subCategory: charge.subCategory,
        mainCategory: charge.mainCategory,
        accountCode: charge.accountCode,
        amount: charge.amount,
        date: header.date,
        remarks: header.remarks
      };

      setLoading(true);
      try {
        const success = await updateCollectionEntry(entry);
        if (success) {
          setItems(prev => prev.map(item => item.id === editingId ? entry : item));
          setNotification({ open: true, message: 'Entry updated successfully', severity: 'success' });
          handleCancelEdit();
        } else {
          throw new Error('Update failed');
        }
      } catch (error) {
        console.error("Error updating entry:", error);
        setNotification({
          open: true,
          message: 'Failed to update entry.',
          severity: 'error'
        });
      } finally {
        setLoading(false);
      }
      return;
    }

    setLoading(true);
    try {
      const nextIdStart = items.length > 0 ? Math.max(...items.map(i => i.id)) + 1 : 1;
      const newRows = preparedCharges.map((c, idx) => ({
        id: nextIdStart + idx,
        afNo: header.afNo,
        orNo: header.orNo,
        payor: header.payor,
        subCategory: c.subCategory,
        mainCategory: c.mainCategory,
        accountCode: c.accountCode,
        amount: c.amount,
        date: header.date,
        remarks: header.remarks
      }));

      const success = await saveCollectionEntryBulk(header, preparedCharges);
      
      if (!success) {
        setNotification({
          open: true,
          message: 'Failed to save to Google Sheet. Please try again.',
          severity: 'error'
        });
        return;
      }

      // Add new items at the beginning
      setItems([...newRows.reverse(), ...items]);
      
      const nextOrNo = getNextOrNo(header.orNo);

      setForm({
        afNo: header.afNo,
        orNo: nextOrNo,
        payor: '',
        date: header.date,
        remarks: ''
      });
      setCharges([{ subCategory: '', mainCategory: '', accountCode: '', amount: '' }]);
      
      setNotification({
        open: true,
        message: 'Entry saved successfully!',
        severity: 'success'
      });

      setTimeout(() => {
        // We use a small timeout to ensure the component is ready
        // For Autocomplete, we might need to find the input element differently if direct ref fails
        // But let's try accessing the ref directly first
        if (payorRef.current) {
          payorRef.current.focus();
        } else {
           // Fallback if ref is not attached to the input directly
           const input = document.querySelector('input[name="payor-input"]');
           (input as HTMLElement)?.focus();
        }
      }, 100);
    } catch (error) {
      console.error("Error saving entry:", error);
      setNotification({
        open: true,
        message: 'An unexpected error occurred while saving.',
        severity: 'error'
      });
    } finally {
      setLoading(false);
    }
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      addItem();
    }
  };

  const totalAmount = filteredItems.reduce((sum, i) => sum + (i.amount || 0), 0);
  const currentEntryTotal = charges.reduce((sum, c) => sum + (Number(c.amount) || 0), 0);

  const clearFilters = () => {
    setSearchTerm('');
    setFilterAfNo(null);
    setFilterSubCategory(null);
    setFilterDate('');
  };

  return (
    <Container maxWidth="xl">
      <Backdrop
        sx={{ color: '#fff', zIndex: (theme) => theme.zIndex.drawer + 1 }}
        open={loading}
      >
        <CircularProgress color="inherit" />
      </Backdrop>

      <Notification
        open={notification.open}
        onClose={handleCloseNotification}
        message={notification.message}
        severity={notification.severity}
      />

      <ConfirmDialog
        open={deleteDialogOpen}
        onClose={() => setDeleteDialogOpen(false)}
        onConfirm={handleDelete}
        title="Delete Entry"
        message="Are you sure you want to delete this entry? This action cannot be undone."
        confirmText="Delete"
        severity="error"
      />

      <Box sx={{ mb: 4, display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <Box>
          <Typography variant="h5" fontWeight="bold">Collections</Typography>
          <Typography variant="body2" color="text.secondary">Manage your collection entries efficiently.</Typography>
        </Box>
        {!showEntryForm && (
          <Button 
            variant="contained" 
            startIcon={<AddCircleOutline />} 
            onClick={() => setShowEntryForm(true)}
          >
            New Entry
          </Button>
        )}
      </Box>

      {showEntryForm && (
      <Card elevation={3} sx={{ mb: 4, borderRadius: 2 }}>
        <CardHeader 
          title={editingId ? `Edit Collection Entry #${editingId}` : "New Collection Entry"}
          subheader={editingId ? "Update the details of the selected entry." : "Fill in the details below to record a new transaction."}
          titleTypographyProps={{ variant: 'h6', fontWeight: 'bold' }}
          action={
            <Button 
              color="inherit" 
              onClick={editingId ? handleCancelEdit : () => setShowEntryForm(false)}
            >
              Cancel
            </Button>
          }
          sx={{ bgcolor: 'grey.50', borderBottom: 1, borderColor: 'divider' }}
        />
        <CardContent sx={{ p: 3 }}>
          <Grid container spacing={4}>
            {/* Left Column: Transaction Details */}
            <Grid size={{ xs: 12, md: 4 }}>
              <Typography variant="subtitle2" color="primary" sx={{ mb: 2, fontWeight: 'bold', textTransform: 'uppercase', letterSpacing: 1 }}>
                Transaction Details
              </Typography>
              
              <Stack spacing={2}>
                <Box sx={{ display: 'flex', gap: 2 }}>
                  <TextField
                    label="AF Number"
                    fullWidth
                    size="small"
                    value={form.afNo}
                    onChange={(e) => setForm({ ...form, afNo: e.target.value })}
                    slotProps={{
                      input: {
                        startAdornment: (
                          <InputAdornment position="start">
                            <ReceiptLong fontSize="small" color="action" />
                          </InputAdornment>
                        ),
                      }
                    }}
                  />
                  <TextField
                    label="OR Number"
                    fullWidth
                    size="small"
                    value={form.orNo}
                    onChange={(e) => {
                      const val = e.target.value;
                      if (val.length <= 8) {
                        setForm({ ...form, orNo: val });
                      }
                    }}
                    onBlur={() => {
                      if (form.orNo && /^\d+$/.test(form.orNo)) {
                        setForm({ ...form, orNo: form.orNo.padStart(8, '0') });
                      }
                    }}
                    slotProps={{
                      input: {
                        startAdornment: (
                          <InputAdornment position="start">
                            <Typography variant="caption" color="text.secondary">#</Typography>
                          </InputAdornment>
                        ),
                      }
                    }}
                  />
                </Box>

                <Autocomplete
                  freeSolo
                  options={uniquePayors}
                  value={form.payor}
                  onChange={(_, newValue) => {
                    setForm({ ...form, payor: newValue || '' });
                  }}
                  onInputChange={(_, newInputValue) => {
                    setForm({ ...form, payor: newInputValue });
                  }}
                  renderInput={(params) => (
                    <TextField
                      {...params}
                      label="Payor Name"
                      fullWidth
                      size="small"
                      name="payor-input"
                      inputRef={(node) => {
                        if (payorRef) {
                          (payorRef as React.MutableRefObject<any>).current = node;
                        }
                        const { ref } = params.InputProps as any || {};
                        if (ref) {
                          if (typeof ref === 'function') ref(node);
                          else ref.current = node;
                        }
                      }}
                      slotProps={{
                        input: {
                          ...params.InputProps,
                          startAdornment: (
                            <>
                              <InputAdornment position="start">
                                <PersonOutline fontSize="small" color="action" />
                              </InputAdornment>
                              {params.InputProps.startAdornment}
                            </>
                          ),
                        }
                      }}
                    />
                  )}
                />

                <TextField
                  type="date"
                  label="Transaction Date"
                  fullWidth
                  size="small"
                  value={form.date}
                  onChange={(e) => setForm({ ...form, date: e.target.value })}
                  InputLabelProps={{ shrink: true }}
                  inputProps={{ tabIndex: -1 }}
                  slotProps={{
                    input: {
                      startAdornment: (
                        <InputAdornment position="start">
                          <Event fontSize="small" color="action" />
                        </InputAdornment>
                      ),
                    }
                  }}
                />

                <TextField
                  label="Remarks"
                  fullWidth
                  multiline
                  rows={3}
                  size="small"
                  value={form.remarks}
                  onChange={(e) => setForm({ ...form, remarks: e.target.value })}
                  inputProps={{ tabIndex: -1 }}
                  slotProps={{
                    input: {
                      startAdornment: (
                        <InputAdornment position="start">
                          <Notes fontSize="small" color="action" />
                        </InputAdornment>
                      ),
                    }
                  }}
                />
              </Stack>
            </Grid>

            {/* Right Column: Charges */}
            <Grid size={{ xs: 12, md: 8 }}>
              <Box sx={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', mb: 2 }}>
                <Typography variant="subtitle2" color="primary" sx={{ fontWeight: 'bold', textTransform: 'uppercase', letterSpacing: 1 }}>
                  Charges
                </Typography>
                <Box sx={{ display: 'flex', alignItems: 'center', gap: 2 }}>
                  <Typography variant="caption" color="text.secondary">
                    {charges.length} item{charges.length !== 1 ? 's' : ''}
                  </Typography>
                  <Typography variant="subtitle2" fontWeight="bold" color="primary">
                    Total: ₱ {currentEntryTotal.toLocaleString('en-PH', { minimumFractionDigits: 2 })}
                  </Typography>
                </Box>
              </Box>

              <Stack spacing={1.5}>
                {charges.map((c, idx) => {
                  const subVal = c.subCategory;
                  const codesForSub = accountCodeOptionsBySub.get(subVal) || [];
                  return (
                    <Paper 
                      key={idx} 
                      variant="outlined" 
                      sx={{ 
                        p: 1.5, 
                        bgcolor: 'grey.50',
                        position: 'relative',
                        transition: 'background-color 0.2s',
                        '&:hover': { bgcolor: 'grey.100' }
                      }}
                    >
                      <Grid container spacing={1} alignItems="center">
                        <Grid size={{ xs: 12, sm: 4 }}>
                          <Autocomplete
                            freeSolo
                            options={subCategories}
                            value={c.subCategory}
                            onChange={(_, v) => {
                              const newVal = v || '';
                              const match = accountCodes.find(ac => ac.subCategory === newVal);
                              const next = [...charges];
                              next[idx] = {
                                subCategory: newVal,
                                mainCategory: match ? match.mainCategory : '',
                                accountCode: match ? match.code : '',
                                amount: c.amount
                              };
                              setCharges(next);
                            }}
                            onInputChange={(_, v) => {
                              const newVal = v;
                              const match = accountCodes.find(ac => ac.subCategory === newVal);
                              const next = [...charges];
                              next[idx] = {
                                subCategory: newVal,
                                mainCategory: match ? match.mainCategory : '',
                                accountCode: match ? match.code : '',
                                amount: c.amount
                              };
                              setCharges(next);
                            }}
                            renderInput={(params) => (
                              <TextField 
                                {...params} 
                                label="Sub Category" 
                                variant="outlined" 
                                size="small" 
                                fullWidth 
                                inputRef={(el) => (subCategoryRefs.current[idx] = el)}
                              />
                            )}
                          />
                        </Grid>
                        <Grid size={{ xs: 6, sm: 3 }}>
                          <Autocomplete
                            freeSolo
                            disabled
                            options={mainCategories}
                            value={c.mainCategory}
                            renderInput={(params) => (
                              <TextField 
                                {...params} 
                                label="Main Category" 
                                variant="outlined" 
                                size="small" 
                                fullWidth 
                              />
                            )}
                          />
                        </Grid>
                        <Grid size={{ xs: 6, sm: 2 }}>
                          <Autocomplete
                            freeSolo
                            disabled
                            options={codesForSub}
                            value={c.accountCode}
                            renderInput={(params) => (
                              <TextField 
                                {...params} 
                                label="Code" 
                                variant="outlined" 
                                size="small" 
                                fullWidth 
                              />
                            )}
                          />
                        </Grid>
                        <Grid size={{ xs: 10, sm: 2 }}>
                          <TextField
                            label="Amount"
                            type="text"
                            fullWidth
                            size="small"
                            value={c.amount}
                            onChange={(e) => {
                              const next = [...charges];
                              next[idx] = { ...next[idx], amount: e.target.value };
                              setCharges(next);
                            }}
                            onBlur={() => {
                              const val = c.amount;
                              if (val && val.trim().startsWith('=')) {
                                try {
                                  const expression = val.trim().substring(1).replace(/\s+/g, '');
                                  const parts = expression.split(/([+\-])/);
                                  let sum = parseFloat(parts[0]) || 0;
                                  for (let i = 1; i < parts.length; i += 2) {
                                    const operator = parts[i];
                                    const operand = parseFloat(parts[i+1]) || 0;
                                    if (operator === '+') sum += operand;
                                    if (operator === '-') sum -= operand;
                                  }
                                  const next = [...charges];
                                  next[idx] = { ...next[idx], amount: String(sum) };
                                  setCharges(next);
                                } catch (e) {
                                  console.error(e);
                                }
                              }
                            }}
                            onKeyDown={(e) => {
                              if (e.key === 'Enter') {
                                const val = c.amount;
                                if (val && val.trim().startsWith('=')) {
                                  e.preventDefault();
                                  try {
                                    const expression = val.trim().substring(1).replace(/\s+/g, '');
                                    const parts = expression.split(/([+\-])/);
                                    let sum = parseFloat(parts[0]) || 0;
                                    for (let i = 1; i < parts.length; i += 2) {
                                      const operator = parts[i];
                                      const operand = parseFloat(parts[i+1]) || 0;
                                      if (operator === '+') sum += operand;
                                      if (operator === '-') sum -= operand;
                                    }
                                    const next = [...charges];
                                    next[idx] = { ...next[idx], amount: String(sum) };
                                    setCharges(next);
                                  } catch (e) {
                                    console.error(e);
                                  }
                                } else {
                                  handleKeyDown(e);
                                }
                              }
                            }}
                            slotProps={{
                              input: {
                                startAdornment: (
                                  <InputAdornment position="start">
                                    <AttachMoney fontSize="small" />
                                  </InputAdornment>
                                ),
                              }
                            }}
                          />
                        </Grid>
                        <Grid size={{ xs: 2, sm: 1 }} sx={{ display: 'flex', justifyContent: 'center' }}>
                           {!editingId && (
                           <Tooltip title="Remove Charge">
                            <IconButton 
                              color="error" 
                              size="small"
                              tabIndex={-1}
                              onClick={() => {
                                const next = charges.filter((_, i) => i !== idx);
                                setCharges(next.length > 0 ? next : [{ subCategory: '', mainCategory: '', accountCode: '', amount: '' }]);
                              }}
                            >
                              <DeleteOutline fontSize="small" />
                            </IconButton>
                          </Tooltip>
                          )}
                        </Grid>
                      </Grid>
                    </Paper>
                  );
                })}
                
                {!editingId && (
                  <Button
                    startIcon={<AddCircleOutline />}
                    onClick={() => {
                      const newIdx = charges.length;
                      setCharges([...charges, { subCategory: '', mainCategory: '', accountCode: '', amount: '' }]);
                      setTimeout(() => {
                        const el = subCategoryRefs.current[newIdx];
                        if (el) {
                          el.focus();
                        }
                      }, 100);
                    }}
                    fullWidth
                    variant="outlined"
                    sx={{ borderStyle: 'dashed', mt: 1 }}
                  >
                    Add Charge
                  </Button>
                )}

                <Button 
                  variant="contained" 
                  color={editingId ? "warning" : "primary"}
                  startIcon={editingId ? <Edit /> : <Save />} 
                  onClick={addItem}
                  fullWidth
                  size="large"
                  disabled={loading}
                  sx={{ mt: 2 }}
                >
                  {loading ? (editingId ? 'Updating...' : 'Saving...') : (editingId ? 'Update Entry' : 'Save Entry')}
                </Button>
              </Stack>
            </Grid>
          </Grid>
        </CardContent>
      </Card>
      )}

      <Paper elevation={2} sx={{ overflow: 'hidden', borderRadius: 2 }}>
        <Box sx={{ p: 2, px: 3, display: 'flex', justifyContent: 'space-between', alignItems: 'center', bgcolor: 'primary.main', color: 'primary.contrastText' }}>
          <Typography variant="h6" fontWeight="bold">Recent Entries</Typography>
          <Typography variant="h6" fontWeight="bold">Total: ₱ {totalAmount.toLocaleString('en-PH', { minimumFractionDigits: 2 })}</Typography>
        </Box>
        
        <Box sx={{ p: 2, bgcolor: 'grey.50', borderBottom: 1, borderColor: 'divider' }}>
          <Grid container spacing={2} alignItems="center">
            <Grid size={{ xs: 12, md: 3 }}>
              <TextField 
                label="Search" 
                placeholder="OR No, Payor, Remarks..."
                fullWidth 
                size="small"
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                slotProps={{
                  input: {
                    startAdornment: (
                      <InputAdornment position="start">
                        <Search fontSize="small" color="action" />
                      </InputAdornment>
                    ),
                  }
                }}
              />
            </Grid>
            <Grid size={{ xs: 12, md: 2 }}>
              <Autocomplete
                options={uniqueAfNos}
                value={filterAfNo}
                onChange={(_, v) => setFilterAfNo(v)}
                renderInput={(params) => <TextField {...params} label="AF No." size="small" />}
              />
            </Grid>
            <Grid size={{ xs: 12, md: 3 }}>
              <Autocomplete
                options={subCategories}
                value={filterSubCategory}
                onChange={(_, v) => setFilterSubCategory(v)}
                renderInput={(params) => <TextField {...params} label="Sub Category" size="small" />}
              />
            </Grid>
            <Grid size={{ xs: 12, md: 2 }}>
              <TextField
                type="month"
                label="Month/Year"
                fullWidth
                size="small"
                value={filterDate}
                onChange={(e) => setFilterDate(e.target.value)}
                InputLabelProps={{ shrink: true }}
              />
            </Grid>
            <Grid size={{ xs: 12, md: 2 }}>
              <Button 
                variant="outlined" 
                color="secondary" 
                startIcon={<Clear />} 
                onClick={clearFilters}
                fullWidth
              >
                Clear
              </Button>
            </Grid>
          </Grid>
        </Box>

        <TableContainer>
          <Table>
            <TableHead sx={{ bgcolor: 'grey.100' }}>
              <TableRow>
                <TableCell sx={{ fontWeight: 'bold' }}></TableCell>
                <TableCell sx={{ fontWeight: 'bold' }}>AF No.</TableCell>
                <TableCell sx={{ fontWeight: 'bold' }}>
                  <TableSortLabel
                    active={sortBy === 'orNo'}
                    direction={sortBy === 'orNo' ? sortOrder : 'asc'}
                    onClick={() => {
                      if (sortBy === 'orNo') {
                        setSortOrder(sortOrder === 'asc' ? 'desc' : 'asc');
                      } else {
                        setSortBy('orNo');
                        setSortOrder('asc');
                      }
                    }}
                  >
                    OR No.
                  </TableSortLabel>
                </TableCell>
                <TableCell sx={{ fontWeight: 'bold' }}>Payor</TableCell>
                <TableCell sx={{ fontWeight: 'bold' }}>
                  <TableSortLabel
                    active={sortBy === 'date'}
                    direction={sortBy === 'date' ? sortOrder : 'asc'}
                    onClick={() => {
                      if (sortBy === 'date') {
                        setSortOrder(sortOrder === 'asc' ? 'desc' : 'asc');
                      } else {
                        setSortBy('date');
                        setSortOrder('desc');
                      }
                    }}
                  >
                    Date
                  </TableSortLabel>
                </TableCell>
                <TableCell align="right" sx={{ fontWeight: 'bold' }}>Total</TableCell>
                <TableCell sx={{ fontWeight: 'bold' }}>Charges</TableCell>
                <TableCell sx={{ fontWeight: 'bold' }}>Remarks</TableCell>
                <TableCell sx={{ fontWeight: 'bold' }}>Actions</TableCell>
              </TableRow>
            </TableHead>
            <TableBody>
              {visibleGroups.map((group) => {
                const expanded = expandedOrNos.includes(group.orNo);
                return (
                  <React.Fragment key={group.orNo}>
                    <TableRow hover>
                      <TableCell>
                        <IconButton size="small" onClick={() => toggleGroup(group.orNo)}>
                          {expanded ? <ExpandLess fontSize="small" /> : <ExpandMore fontSize="small" />}
                        </IconButton>
                      </TableCell>
                      <TableCell>{group.afNo}</TableCell>
                      <TableCell>{group.orNo}</TableCell>
                      <TableCell>{group.payor}</TableCell>
                      <TableCell>{formatDisplayDate(group.date)}</TableCell>
                      <TableCell align="right" sx={{ fontWeight: 'bold', color: 'success.main' }}>
                        ₱ {group.total.toLocaleString('en-PH', { minimumFractionDigits: 2 })}
                      </TableCell>
                      <TableCell>{group.charges.length} charge{group.charges.length !== 1 ? 's' : ''}</TableCell>
                      <TableCell>{group.remarks}</TableCell>
                      <TableCell>
                        <Tooltip title={expanded ? 'Collapse charges' : 'Expand charges'}>
                          <IconButton size="small" onClick={() => toggleGroup(group.orNo)}>
                            {expanded ? <ExpandLess fontSize="small" /> : <ExpandMore fontSize="small" />}
                          </IconButton>
                        </Tooltip>
                      </TableCell>
                    </TableRow>
                    {expanded && (
                      <TableRow>
                        <TableCell colSpan={9} sx={{ bgcolor: 'grey.50', p: 0 }}>
                          <Table size="small">
                            <TableHead>
                              <TableRow>
                                <TableCell sx={{ fontWeight: 'bold' }}>Sub Category</TableCell>
                                <TableCell sx={{ fontWeight: 'bold' }}>Main Category</TableCell>
                                <TableCell sx={{ fontWeight: 'bold' }}>Account Code</TableCell>
                                <TableCell align="right" sx={{ fontWeight: 'bold' }}>Amount</TableCell>
                                <TableCell sx={{ fontWeight: 'bold' }}>Remarks</TableCell>
                                <TableCell sx={{ fontWeight: 'bold' }}>Actions</TableCell>
                              </TableRow>
                            </TableHead>
                            <TableBody>
                              {group.charges.map((charge) => (
                                <TableRow key={charge.id} hover>
                                  <TableCell>{charge.subCategory}</TableCell>
                                  <TableCell>{charge.mainCategory}</TableCell>
                                  <TableCell sx={{ fontFamily: 'monospace', bgcolor: 'grey.100', px: 1, borderRadius: 1 }}>{charge.accountCode}</TableCell>
                                  <TableCell align="right" sx={{ fontWeight: 'bold', color: 'success.main' }}>
                                    ₱ {charge.amount.toLocaleString('en-PH', { minimumFractionDigits: 2 })}
                                  </TableCell>
                                  <TableCell>{charge.remarks}</TableCell>
                                  <TableCell>
                                    <Stack direction="row" spacing={0}>
                                      <Tooltip title="Edit">
                                        <IconButton size="small" color="primary" onClick={() => handleEdit(charge)}>
                                          <Edit fontSize="small" />
                                        </IconButton>
                                      </Tooltip>
                                      <Tooltip title="Delete">
                                        <IconButton size="small" color="error" onClick={() => confirmDelete(charge.id)}>
                                          <DeleteOutline fontSize="small" />
                                        </IconButton>
                                      </Tooltip>
                                    </Stack>
                                  </TableCell>
                                </TableRow>
                              ))}
                            </TableBody>
                          </Table>
                        </TableCell>
                      </TableRow>
                    )}
                  </React.Fragment>
                );
              })}
              {groupedItems.length === 0 && (
                <TableRow>
                  <TableCell colSpan={9} align="center" sx={{ py: 4 }}>
                    <Typography color="text.secondary" variant="body1">
                      {loading ? 'Loading entries...' : 'No entries found.'}
                    </Typography>
                  </TableCell>
                </TableRow>
              )}
            </TableBody>
          </Table>
        </TableContainer>
        <TablePagination
          rowsPerPageOptions={[10, 25, 50]}
          component="div"
          count={filteredItems.length}
          rowsPerPage={rowsPerPage}
          page={page}
          onPageChange={handleChangePage}
          onRowsPerPageChange={handleChangeRowsPerPage}
          labelRowsPerPage="Entries per page:"
        />
      </Paper>
    </Container>
  );
};
