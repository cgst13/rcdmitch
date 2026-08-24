import React, { useState, useEffect, useRef } from 'react';
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
  IconButton, 
  Dialog,
  DialogTitle,
  DialogContent,
  DialogActions,
  TextField,
  Grid,
  CircularProgress,
  Backdrop,
  Divider,
  Card,
  CardContent,
  CardHeader,
  Stack,
  TablePagination,
  InputAdornment,
  Autocomplete
} from '@mui/material';
import { Add, Edit, Delete, Refresh, Search, Clear } from '@mui/icons-material';
import { ConfirmDialog } from '../components/ConfirmDialog';
import { getRPTCollections, saveRPTCollection, deleteRPTCollection } from '../services/googleSheets';
import type { RPTCollectionItem } from '../types/rcd';

// Date Format Helpers
 const formatToMmDdYyyy = (dateStr: string | undefined): string => {
    if (!dateStr) return '';
    
    const s = String(dateStr).trim();
    
    // 1. Handle ISO strings with 'T' - these are likely shifted UTC dates
    if (s.includes('T')) {
      const d = new Date(s);
      if (!isNaN(d.getTime())) {
        const month = String(d.getMonth() + 1).padStart(2, '0');
        const day = String(d.getDate()).padStart(2, '0');
        const year = d.getFullYear();
        return `${month}/${day}/${year}`;
      }
    }

    // 2. Handle simple date strings (remove time part if any)
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
  
  const parseDateToInputFormat = (dateStr: string | undefined): string => {
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

export const RPTCollectionPage: React.FC = () => {
  const [collections, setCollections] = useState<RPTCollectionItem[]>([]);
  const [loading, setLoading] = useState(false);
  const [showEntryForm, setShowEntryForm] = useState(false);
  const [isEditing, setIsEditing] = useState(false);
  const [currentId, setCurrentId] = useState<number | null>(null);

  // Form Ref for focusing
  const payorInputRef = useRef<HTMLInputElement>(null);

  // Pagination State
  const [page, setPage] = useState(0);
  const [rowsPerPage, setRowsPerPage] = useState(10);

  // Filter State
  const [searchTerm, setSearchTerm] = useState('');
  const [filterDate, setFilterDate] = useState('');
  const [filterAF56Id, setFilterAF56Id] = useState<string | null>(null);

  // Derived State for Filters
  const uniqueAF56Ids = Array.from(new Set(collections.map(c => c.af56Id).filter(Boolean))).sort();

  // Suggestions State
  const payorSuggestions = Array.from(new Set(collections.map(c => c.payor).filter(Boolean))).sort();
  const barangaySuggestions = Array.from(new Set(collections.map(c => c.barangay).filter(Boolean))).sort();
  const landNameSuggestions = Array.from(new Set(collections.map(c => c.landName).filter(Boolean))).sort();

  // Delete Dialog State
  const [deleteDialogOpen, setDeleteDialogOpen] = useState(false);
  const [itemToDelete, setItemToDelete] = useState<number | null>(null);

  // Details Dialog State
  const [viewDialogOpen, setViewDialogOpen] = useState(false);
  const [selectedItem, setSelectedItem] = useState<RPTCollectionItem | null>(null);

  // Form State
  const [formData, setFormData] = useState({
    af56Id: '',
    orNumber: '',
    payor: '',
    barangay: '',
    landName: '',
    tdNumber: '',
    yearsPaid: '',
    amount: 0,
    date: new Date().toISOString().split('T')[0],
    remarks: ''
  });

  const loadCollections = async () => {
    setLoading(true);
    try {
      const data = await getRPTCollections();
      setCollections(data);
    } catch (error) {
      console.error('Failed to load collections', error);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    loadCollections();
  }, []);

  const handleOpen = () => {
    // Autofill from latest record
    let latestRecord: RPTCollectionItem | null = null;
    if (collections.length > 0) {
      // Sort by date and then by ID to find the latest
      const sorted = [...collections].sort((a, b) => {
        const s1 = parseDateToInputFormat(a.date);
        const s2 = parseDateToInputFormat(b.date);
        if (s2 !== s1) return s2.localeCompare(s1);
        return b.id - a.id;
      });
      latestRecord = sorted[0];
    }

    // Calculate next OR number
    let nextOr = '';
    if (latestRecord?.orNumber) {
      const currentOr = parseInt(latestRecord.orNumber, 10);
      if (!isNaN(currentOr)) {
        nextOr = String(currentOr + 1).padStart(8, '0');
      }
    }

    setFormData({
      af56Id: latestRecord?.af56Id || '',
      orNumber: nextOr, // Use incremented OR with padding
      payor: '',
      barangay: '',
      landName: '',
      tdNumber: '',
      yearsPaid: '',
      amount: 0,
      date: parseDateToInputFormat(latestRecord?.date),
      remarks: ''
    });
    setIsEditing(false);
    setShowEntryForm(true);
    
    // Small delay to ensure form is rendered before focusing
    setTimeout(() => {
      if (payorInputRef.current) {
        payorInputRef.current.focus();
      }
    }, 100);
  };

  const handleEdit = (item: RPTCollectionItem) => {
    setFormData({
      af56Id: item.af56Id,
      orNumber: item.orNumber,
      payor: item.payor,
      barangay: item.barangay,
      landName: item.landName,
      tdNumber: item.tdNumber,
      yearsPaid: item.yearsPaid,
      amount: item.amount,
      date: parseDateToInputFormat(item.date),
      remarks: item.remarks || ''
    });
    setCurrentId(item.id);
    setIsEditing(true);
    setShowEntryForm(true);
    // Scroll to top to see the form
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  const handleCancelEdit = () => {
    setShowEntryForm(false);
    setIsEditing(false);
    setCurrentId(null);
  };

  const handleDeleteClick = (id: number) => {
    setItemToDelete(id);
    setDeleteDialogOpen(true);
  };

  const handleConfirmDelete = async () => {
    if (itemToDelete) {
        setLoading(true);
        try {
            await deleteRPTCollection(itemToDelete);
            await loadCollections();
        } catch (error) {
            console.error('Failed to delete collection', error);
        } finally {
            setLoading(false);
            setDeleteDialogOpen(false);
            setItemToDelete(null);
        }
    }
  };

  const handleSave = async () => {
    setLoading(true);
    try {
      // Calculate next ID
      let nextId = Date.now();
      if (!isEditing) {
        const maxId = collections.reduce((max, c) => Math.max(max, c.id), 0);
        nextId = maxId + 1;
      }

      const collection: RPTCollectionItem = {
        id: isEditing && currentId ? currentId : nextId,
        ...formData,
        date: formatToMmDdYyyy(formData.date) // Store in sheet as mm/dd/yyyy
      };
      
      await saveRPTCollection(collection);
      await loadCollections();
      
      if (isEditing) {
        setShowEntryForm(false);
        setIsEditing(false);
        setCurrentId(null);
      } else {
        // For new entries, keep the form open but reset and focus Payor
        // Increment OR Number automatically if it's numeric
        let nextOrNumber = formData.orNumber;
        const orNum = parseInt(formData.orNumber, 10);
        if (!isNaN(orNum)) {
          nextOrNumber = String(orNum + 1).padStart(8, '0'); // Always 8 digits
        }

        setFormData({
          ...formData,
          orNumber: nextOrNumber,
          payor: '',
          barangay: '',
          landName: '',
          tdNumber: '',
          yearsPaid: '',
          amount: 0,
          remarks: ''
        });
        
        // Refocus the payor field
        setTimeout(() => {
          if (payorInputRef.current) {
            payorInputRef.current.focus();
          }
        }, 100);
      }
    } catch (error) {
      console.error('Failed to save collection', error);
    } finally {
      setLoading(false);
    }
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter' && isFormValid && !loading) {
      e.preventDefault();
      handleSave();
    }
  };

  // Pagination Handlers
  const handleChangePage = (_event: unknown, newPage: number) => {
    setPage(newPage);
  };

  const handleChangeRowsPerPage = (event: React.ChangeEvent<HTMLInputElement>) => {
    setRowsPerPage(parseInt(event.target.value, 10));
    setPage(0);
  };

  // Filter Logic
  const filteredCollections = collections.filter(item => {
    const matchesSearch = 
      (item.orNumber?.toLowerCase().includes(searchTerm.toLowerCase()) || '') ||
      (item.payor?.toLowerCase().includes(searchTerm.toLowerCase()) || '') ||
      (item.remarks?.toLowerCase().includes(searchTerm.toLowerCase()) || '') ||
      (item.landName?.toLowerCase().includes(searchTerm.toLowerCase()) || '') ||
      (item.barangay?.toLowerCase().includes(searchTerm.toLowerCase()) || '');
    
    // Normalize date for month/year filtering
    const normalizedDate = parseDateToInputFormat(item.date);
    const matchesDate = filterDate ? normalizedDate.startsWith(filterDate) : true;
    const matchesAF56Id = filterAF56Id ? item.af56Id === filterAF56Id : true;

    return matchesSearch && matchesDate && matchesAF56Id;
  }).sort((a, b) => {
    const s1 = parseDateToInputFormat(a.date);
    const s2 = parseDateToInputFormat(b.date);
    if (s2 !== s1) return s2.localeCompare(s1);
    return b.id - a.id;
  });

  const paginatedCollections = filteredCollections.slice(
    page * rowsPerPage,
    page * rowsPerPage + rowsPerPage
  );

  const totalAmount = filteredCollections.reduce((sum, item) => sum + (item.amount || 0), 0);

  const isCancelled = formData.payor.trim().toUpperCase() === 'CANCELLED';
  const isFormValid = formData.orNumber && formData.payor && (isCancelled || formData.amount > 0);

  const clearFilters = () => {
    setSearchTerm('');
    setFilterDate('');
    setFilterAF56Id(null);
  };

  return (
    <Box sx={{ p: 3 }}>
        <Backdrop open={loading} sx={{ color: '#fff', zIndex: (theme) => theme.zIndex.drawer + 1 }}>
            <CircularProgress color="inherit" />
        </Backdrop>

        <Box display="flex" justifyContent="space-between" alignItems="center" mb={3}>
            <Typography variant="h4" component="h1">
                RPT Collection
            </Typography>
            <Box>
                <Button 
                    startIcon={<Refresh />} 
                    onClick={loadCollections} 
                    sx={{ mr: 1 }}
                >
                    Refresh
                </Button>
                {!showEntryForm && (
                  <Button 
                      variant="contained" 
                      startIcon={<Add />} 
                      onClick={handleOpen}
                  >
                      Add Collection
                  </Button>
                )}
            </Box>
        </Box>

        {showEntryForm && (
          <Card elevation={3} sx={{ mb: 4, borderRadius: 2 }}>
            <CardHeader 
              title={isEditing ? `Edit Collection Entry #${currentId}` : "New Collection Entry"}
              subheader={isEditing ? "Update the details of the selected entry." : "Fill in the details below to record a new transaction."}
              titleTypographyProps={{ variant: 'h6', fontWeight: 'bold' }}
              action={
                <Button 
                  color="inherit" 
                  onClick={handleCancelEdit}
                >
                  Cancel
                </Button>
              }
              sx={{ bgcolor: 'grey.50', borderBottom: 1, borderColor: 'divider' }}
            />
            <CardContent>
              <Grid container spacing={2}>
                <Grid size={{ xs: 12, sm: 6 }}>
                    <TextField
                        label="Date"
                        type="date"
                        fullWidth
                        value={formData.date}
                        onChange={(e) => setFormData({ ...formData, date: e.target.value })}
                        onKeyDown={handleKeyDown}
                        required
                        InputLabelProps={{ shrink: true }}
                    />
                </Grid>
                <Grid size={{ xs: 12, sm: 6 }}>
                    <TextField
                        label="AF56 ID"
                        fullWidth
                        value={formData.af56Id}
                        onChange={(e) => setFormData({ ...formData, af56Id: e.target.value })}
                        onKeyDown={handleKeyDown}
                        required
                    />
                </Grid>
                <Grid size={{ xs: 12, sm: 6 }}>
                    <TextField
                        label="OR Number"
                        fullWidth
                        value={formData.orNumber}
                        onChange={(e) => setFormData({ ...formData, orNumber: e.target.value })}
                        onKeyDown={handleKeyDown}
                        required
                    />
                </Grid>
                <Grid size={{ xs: 12, sm: 6 }}>
                    <Autocomplete
                        freeSolo
                        options={payorSuggestions}
                        value={formData.payor}
                        onInputChange={(_, newValue) => setFormData({ ...formData, payor: newValue })}
                        onChange={(_, newValue) => setFormData({ ...formData, payor: newValue || '' })}
                        renderInput={(params) => (
                            <TextField
                                {...params}
                                inputRef={payorInputRef}
                                label="Payor"
                                onKeyDown={handleKeyDown}
                                required
                                fullWidth
                            />
                        )}
                    />
                </Grid>
                <Grid size={{ xs: 12, sm: 6 }}>
                    <Autocomplete
                        freeSolo
                        options={barangaySuggestions}
                        value={formData.barangay}
                        onInputChange={(_, newValue) => setFormData({ ...formData, barangay: newValue })}
                        onChange={(_, newValue) => setFormData({ ...formData, barangay: newValue || '' })}
                        renderInput={(params) => (
                            <TextField
                                {...params}
                                label="Barangay"
                                onKeyDown={handleKeyDown}
                                fullWidth
                            />
                        )}
                    />
                </Grid>
                <Grid size={{ xs: 12, sm: 6 }}>
                    <Autocomplete
                        freeSolo
                        options={landNameSuggestions}
                        value={formData.landName}
                        onInputChange={(_, newValue) => setFormData({ ...formData, landName: newValue })}
                        onChange={(_, newValue) => setFormData({ ...formData, landName: newValue || '' })}
                        renderInput={(params) => (
                            <TextField
                                {...params}
                                label="Land Name"
                                onKeyDown={handleKeyDown}
                                fullWidth
                            />
                        )}
                    />
                </Grid>
                <Grid size={{ xs: 12, sm: 6 }}>
                    <TextField
                        label="TD Number"
                        fullWidth
                        value={formData.tdNumber}
                        onChange={(e) => setFormData({ ...formData, tdNumber: e.target.value })}
                        onKeyDown={handleKeyDown}
                    />
                </Grid>
                <Grid size={{ xs: 12, sm: 6 }}>
                    <TextField
                        label="Years Paid"
                        fullWidth
                        value={formData.yearsPaid}
                        onChange={(e) => setFormData({ ...formData, yearsPaid: e.target.value })}
                        onKeyDown={handleKeyDown}
                    />
                </Grid>
                <Grid size={{ xs: 12, sm: 6 }}>
                    <TextField
                        label="Amount"
                        type="number"
                        fullWidth
                        value={formData.amount}
                        onChange={(e) => setFormData({ ...formData, amount: parseFloat(e.target.value) || 0 })}
                        onKeyDown={handleKeyDown}
                        required
                    />
                </Grid>
                <Grid size={{ xs: 12 }}>
                    <TextField
                        label="Remarks"
                        fullWidth
                        multiline
                        rows={2}
                        value={formData.remarks}
                        onChange={(e) => setFormData({ ...formData, remarks: e.target.value })}
                        onKeyDown={(e) => {
                          // Allow newline in remarks with Shift+Enter, but save on Enter
                          if (e.key === 'Enter' && !e.shiftKey) {
                            handleKeyDown(e);
                          }
                        }}
                    />
                </Grid>
                <Grid size={{ xs: 12 }}>
                  <Stack direction="row" spacing={2} justifyContent="flex-end" sx={{ mt: 2 }}>
                    <Button onClick={handleCancelEdit}>Cancel</Button>
                    <Button 
                      onClick={handleSave} 
                      variant="contained" 
                      disabled={!isFormValid}
                    >
                      {isEditing ? 'Update Entry' : 'Save Entry'}
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
                <Grid size={{ xs: 12, md: 4 }}>
                  <TextField 
                    label="Search" 
                    placeholder="Search..."
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
                <Grid size={{ xs: 12, md: 3 }}>
                  <Autocomplete
                    options={uniqueAF56Ids}
                    value={filterAF56Id}
                    onChange={(_, newValue) => setFilterAF56Id(newValue)}
                    renderInput={(params) => <TextField {...params} label="AF56 ID" size="small" />}
                  />
                </Grid>
                <Grid size={{ xs: 12, md: 3 }}>
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
                    <TableHead>
                        <TableRow>
                            <TableCell sx={{ fontWeight: 'bold' }}>AF56 ID</TableCell>
                            <TableCell sx={{ fontWeight: 'bold' }}>OR Number</TableCell>
                            <TableCell sx={{ fontWeight: 'bold' }}>Date</TableCell>
                            <TableCell sx={{ fontWeight: 'bold' }}>Barangay</TableCell>
                            <TableCell sx={{ fontWeight: 'bold' }}>Land Name</TableCell>
                            <TableCell sx={{ fontWeight: 'bold', textAlign: 'right' }}>Amount</TableCell>
                            <TableCell sx={{ fontWeight: 'bold' }}>Remarks</TableCell>
                            <TableCell sx={{ fontWeight: 'bold', textAlign: 'right' }}>Actions</TableCell>
                        </TableRow>
                    </TableHead>
                    <TableBody>
                        {paginatedCollections.map((row) => (
                            <TableRow 
                                key={row.id} 
                                hover
                                onClick={() => { setSelectedItem(row); setViewDialogOpen(true); }}
                                sx={{ cursor: 'pointer' }}
                            >
                                <TableCell>{row.af56Id}</TableCell>
                                <TableCell>{row.orNumber}</TableCell>
                                <TableCell>{formatToMmDdYyyy(row.date)}</TableCell>
                                <TableCell>{row.barangay}</TableCell>
                                <TableCell>{row.landName}</TableCell>
                                <TableCell align="right">₱ {(row.amount || 0).toLocaleString('en-PH', { minimumFractionDigits: 2 })}</TableCell>
                                <TableCell>{row.remarks}</TableCell>
                                <TableCell align="right">
                                    <IconButton 
                                        color="primary" 
                                        onClick={(e) => { e.stopPropagation(); handleEdit(row); }}
                                    >
                                        <Edit />
                                    </IconButton>
                                    <IconButton 
                                        color="error" 
                                        onClick={(e) => { e.stopPropagation(); handleDeleteClick(row.id); }}
                                    >
                                        <Delete />
                                    </IconButton>
                                </TableCell>
                            </TableRow>
                        ))}
                        {paginatedCollections.length === 0 && (
                            <TableRow>
                                <TableCell colSpan={7} align="center">
                                    No RPT collections found
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
        </Paper>

        <Dialog 
            open={viewDialogOpen} 
            onClose={() => setViewDialogOpen(false)} 
            maxWidth="sm" 
            fullWidth
        >
            <DialogTitle sx={{ bgcolor: 'primary.main', color: 'white', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <Typography variant="h6">RPT Collection Details</Typography>
                <Typography variant="subtitle2" sx={{ opacity: 0.8 }}>ID: {selectedItem?.id}</Typography>
            </DialogTitle>
            <DialogContent>
                {selectedItem && (
                    <Box sx={{ mt: 3 }}>
                        <Grid container spacing={3}>
                            {/* Transaction Info */}
                            <Grid size={{ xs: 12 }}>
                                <Card variant="outlined" sx={{ bgcolor: '#f8f9fa' }}>
                                    <CardContent sx={{ py: 2 }}>
                                        <Grid container spacing={2}>
                                            <Grid size={{ xs: 6 }}>
                                                <Typography variant="caption" color="text.secondary">OR Number</Typography>
                                                <Typography variant="h6" color="primary" sx={{ fontWeight: 'bold' }}>
                                                    {selectedItem.orNumber}
                                                </Typography>
                                            </Grid>
                                            <Grid size={{ xs: 6 }} sx={{ textAlign: 'right' }}>
                                                <Typography variant="caption" color="text.secondary">Date</Typography>
                                                <Typography variant="body1">
                                                    {formatToMmDdYyyy(selectedItem.date)}
                                                </Typography>
                                            </Grid>
                                        </Grid>
                                    </CardContent>
                                </Card>
                            </Grid>

                            {/* Payor Info */}
                            <Grid size={{ xs: 12 }}>
                                <Typography variant="caption" color="text.secondary">Payor Name</Typography>
                                <Typography variant="h5" sx={{ fontWeight: 500 }}>
                                    {selectedItem.payor}
                                </Typography>
                                <Divider sx={{ my: 1 }} />
                            </Grid>

                            {/* Property Details */}
                            <Grid size={{ xs: 6 }}>
                                <Typography variant="caption" color="text.secondary">Land Name</Typography>
                                <Typography variant="body1">{selectedItem.landName || '-'}</Typography>
                            </Grid>
                            <Grid size={{ xs: 6 }}>
                                <Typography variant="caption" color="text.secondary">Barangay</Typography>
                                <Typography variant="body1">{selectedItem.barangay || '-'}</Typography>
                            </Grid>
                            <Grid size={{ xs: 6 }}>
                                <Typography variant="caption" color="text.secondary">TD Number</Typography>
                                <Typography variant="body1">{selectedItem.tdNumber || '-'}</Typography>
                            </Grid>
                            <Grid size={{ xs: 6 }}>
                                <Typography variant="caption" color="text.secondary">AF56 ID</Typography>
                                <Typography variant="body1">{selectedItem.af56Id || '-'}</Typography>
                            </Grid>

                            <Grid size={{ xs: 12 }}>
                                <Divider />
                            </Grid>

                            {/* Payment Info */}
                            <Grid size={{ xs: 6 }}>
                                <Typography variant="caption" color="text.secondary">Years Paid</Typography>
                                <Typography variant="body1" sx={{ fontWeight: 'bold' }}>{selectedItem.yearsPaid || '-'}</Typography>
                            </Grid>
                            <Grid size={{ xs: 6 }} sx={{ textAlign: 'right' }}>
                                <Typography variant="caption" color="text.secondary">Total Amount</Typography>
                                <Typography variant="h5" color="success.main" sx={{ fontWeight: 'bold' }}>
                                    ₱ {(selectedItem.amount || 0).toLocaleString('en-PH', { minimumFractionDigits: 2 })}
                                </Typography>
                            </Grid>

                            {/* Remarks */}
                            {selectedItem.remarks && (
                                <Grid size={{ xs: 12 }}>
                                    <Box sx={{ bgcolor: '#fff3e0', p: 2, borderRadius: 1 }}>
                                        <Typography variant="caption" color="text.secondary">Remarks</Typography>
                                        <Typography variant="body2" sx={{ fontStyle: 'italic' }}>
                                            {selectedItem.remarks}
                                        </Typography>
                                    </Box>
                                </Grid>
                            )}
                        </Grid>
                    </Box>
                )}
            </DialogContent>
            <DialogActions sx={{ p: 2 }}>
                <Button 
                    onClick={() => { setViewDialogOpen(false); if(selectedItem) handleEdit(selectedItem); }}
                    startIcon={<Edit />}
                >
                    Edit
                </Button>
                <Button onClick={() => setViewDialogOpen(false)} variant="contained">
                    Close
                </Button>
            </DialogActions>
        </Dialog>

        <ConfirmDialog
            open={deleteDialogOpen}
            onClose={() => setDeleteDialogOpen(false)}
            onConfirm={handleConfirmDelete}
            title="Delete RPT Collection"
            message="Are you sure you want to delete this RPT collection? This action cannot be undone."
            confirmText="Delete"
            severity="error"
        />
    </Box>
  );
};
