import React, { useState, useEffect } from 'react';
import { 
  Box, 
  Typography, 
  Paper, 
  List, 
  ListItem, 
  ListItemText, 
  ListItemIcon, 
  ListItemButton,
  Switch, 
  Divider,
  Avatar,
  Button,
  TextField,
  InputAdornment,
  IconButton,
  Chip,
  CircularProgress,
  Alert,
  Tooltip
} from '@mui/material';
import { 
  Notifications, 
  DarkMode, 
  Language, 
  Security, 
  Person, 
  CloudDownload,
  DeleteForever,
  Info,
  TableChart,
  Link as LinkIcon,
  OpenInNew,
  ContentPaste,
  ContentCopy,
  Save,
  CheckCircle,
  Refresh,
  Storage
} from '@mui/icons-material';
import { useAuth } from '../context/useAuth';
import { ConfirmDialog } from '../components/ConfirmDialog';
import { Notification } from '../components/Notification';
import { 
  getSpreadsheetConfig, 
  updateSpreadsheetConfig, 
  extractSpreadsheetId,
  getApiUrl,
  setApiUrl
} from '../services/googleSheets';

const SERVICE_ACCOUNT_EMAIL = 'credentials@rcd-lguconcepcion.iam.gserviceaccount.com';

export const SettingsPage: React.FC = () => {
  const { user } = useAuth();
  const [openClearDialog, setOpenClearDialog] = useState(false);
  const [notification, setNotification] = useState<{
    open: boolean;
    message: string;
    severity: 'success' | 'error' | 'info' | 'warning';
  }>({ open: false, message: '', severity: 'info' });

  // Google Sheet Database state
  const [sheetInput, setSheetInput] = useState('');
  const [activeSheetId, setActiveSheetId] = useState('');
  const [activeSheetUrl, setActiveSheetUrl] = useState('');
  const [apiUrlInput, setApiUrlInput] = useState(getApiUrl());
  const [isLoadingConfig, setIsLoadingConfig] = useState(true);
  const [isSavingSheet, setIsSavingSheet] = useState(false);
  const [isTestingConn, setIsTestingConn] = useState(false);
  const [connectionStatus, setConnectionStatus] = useState<'connected' | 'error' | 'checking'>('checking');
  const [errorMessage, setErrorMessage] = useState<string | null>(null);

  // Load current spreadsheet config on mount
  useEffect(() => {
    loadDatabaseConfig();
  }, []);

  const loadDatabaseConfig = async () => {
    setIsLoadingConfig(true);
    setConnectionStatus('checking');
    setErrorMessage(null);
    try {
      const config = await getSpreadsheetConfig();
      setActiveSheetId(config.spreadsheetId);
      setActiveSheetUrl(config.spreadsheetUrl);
      setSheetInput(config.spreadsheetUrl);
      setConnectionStatus('connected');
    } catch (err) {
      console.error('Failed to load database config:', err);
      setConnectionStatus('error');
    } finally {
      setIsLoadingConfig(false);
    }
  };

  const handleSaveSpreadsheet = async () => {
    if (!sheetInput.trim()) {
      setNotification({
        open: true,
        message: 'Please paste a valid Google Sheet link or ID.',
        severity: 'warning'
      });
      return;
    }

    const extractedId = extractSpreadsheetId(sheetInput);
    if (!extractedId) {
      setNotification({
        open: true,
        message: 'Invalid Google Sheet link or ID format. Please double check the link.',
        severity: 'error'
      });
      return;
    }

    setIsSavingSheet(true);
    setConnectionStatus('checking');
    setErrorMessage(null);

    try {
      const res = await updateSpreadsheetConfig(sheetInput);
      if (res.success && res.config) {
        setActiveSheetId(res.config.spreadsheetId);
        setActiveSheetUrl(res.config.spreadsheetUrl);
        setSheetInput(res.config.spreadsheetUrl);
        setConnectionStatus('connected');
        setNotification({
          open: true,
          message: res.message || 'Google Sheet database link updated successfully!',
          severity: 'success'
        });
      } else {
        setConnectionStatus('error');
        setErrorMessage(res.message || 'Failed to update Google Sheet link.');
        setNotification({
          open: true,
          message: res.message || 'Failed to update Google Sheet link.',
          severity: 'error'
        });
      }
    } catch (err: any) {
      setConnectionStatus('error');
      const msg = err.message || 'An error occurred while saving the Google Sheet database link.';
      setErrorMessage(msg);
      setNotification({
        open: true,
        message: msg,
        severity: 'error'
      });
    } finally {
      setIsSavingSheet(false);
    }
  };

  const handleTestConnection = async () => {
    setIsTestingConn(true);
    setErrorMessage(null);
    try {
      const config = await getSpreadsheetConfig();
      if (config.spreadsheetId) {
        setConnectionStatus('connected');
        setNotification({
          open: true,
          message: `Database connection verified! Sheet ID: ${config.spreadsheetId}`,
          severity: 'success'
        });
      } else {
        setConnectionStatus('error');
        setNotification({
          open: true,
          message: 'Unable to connect to Google Sheet database.',
          severity: 'error'
        });
      }
    } catch {
      setConnectionStatus('error');
      setNotification({
        open: true,
        message: 'Database connection check failed.',
        severity: 'error'
      });
    } finally {
      setIsTestingConn(false);
    }
  };

  const handlePasteClipboard = async () => {
    try {
      const text = await navigator.clipboard.readText();
      if (text) {
        setSheetInput(text.trim());
        setNotification({
          open: true,
          message: 'Pasted link from clipboard.',
          severity: 'info'
        });
      }
    } catch {
      setNotification({
        open: true,
        message: 'Could not read clipboard. Please paste manually.',
        severity: 'warning'
      });
    }
  };

  const handleCopyServiceAccount = () => {
    navigator.clipboard.writeText(SERVICE_ACCOUNT_EMAIL);
    setNotification({
      open: true,
      message: 'Service Account email copied to clipboard!',
      severity: 'success'
    });
  };

  const handleExportData = () => {
    const data = localStorage.getItem('rcd_reports');
    if (!data) {
        setNotification({ open: true, message: 'No data to export.', severity: 'warning' });
        return;
    }
    
    const blob = new Blob([data], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `rcd_backup_${new Date().toISOString().split('T')[0]}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setNotification({ open: true, message: 'Data exported successfully.', severity: 'success' });
  };

  const handleClearData = () => {
    localStorage.removeItem('rcd_reports');
    setOpenClearDialog(false);
    setNotification({ open: true, message: 'Local data cleared successfully.', severity: 'success' });
  };

  return (
    <Box>
      <Typography variant="h4" fontWeight="bold" gutterBottom sx={{ mb: 4 }}>
        Settings
      </Typography>

      {/* Profile Section */}
      <Paper sx={{ p: 3, mb: 3, display: 'flex', alignItems: 'center', gap: 3, borderRadius: 2 }}>
        <Avatar 
            sx={{ width: 64, height: 64, bgcolor: 'primary.main', fontSize: '2rem' }}
        >
            {user?.name?.charAt(0) || 'U'}
        </Avatar>
        <Box sx={{ flex: 1 }}>
            <Typography variant="h6" fontWeight="bold">{user?.name || 'User'}</Typography>
            <Typography variant="body2" color="text.secondary">{user?.email || 'user@example.com'}</Typography>
            <Typography variant="caption" sx={{ display: 'block', mt: 0.5, color: 'primary.main', fontWeight: 'medium' }}>
                {user?.role || 'Administrator'}
            </Typography>
        </Box>
        <Button variant="outlined" startIcon={<Person />}>
            Edit Profile
        </Button>
      </Paper>

      {/* Google Sheet Database Section */}
      <Paper sx={{ p: 3, mb: 3, borderRadius: 2, border: '1px solid', borderColor: 'divider' }}>
        <Box sx={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', mb: 2 }}>
          <Box sx={{ display: 'flex', alignItems: 'center', gap: 1.5 }}>
            <Avatar sx={{ bgcolor: '#0f9d58', width: 40, height: 40 }}>
              <TableChart />
            </Avatar>
            <Box>
              <Typography variant="h6" fontWeight="bold">
                Google Sheet Database
              </Typography>
              <Typography variant="body2" color="text.secondary">
                Configure the primary Google Sheet used as the database for this system
              </Typography>
            </Box>
          </Box>

          {connectionStatus === 'connected' && (
            <Chip 
              icon={<CheckCircle />} 
              label="Connected" 
              color="success" 
              variant="outlined" 
              size="small" 
            />
          )}
          {connectionStatus === 'checking' && (
            <Chip 
              icon={<CircularProgress size={14} color="inherit" />} 
              label="Checking..." 
              color="info" 
              variant="outlined" 
              size="small" 
            />
          )}
          {connectionStatus === 'error' && (
            <Chip 
              label="Connection Issue" 
              color="error" 
              variant="outlined" 
              size="small" 
            />
          )}
        </Box>

        <Divider sx={{ my: 2 }} />

        {/* Service Account Permission Card */}
        <Box sx={{ mb: 3, p: 2, bgcolor: '#f0f4f9', borderRadius: 1.5, border: '1px solid #d3e3fd' }}>
          <Typography variant="caption" color="primary.main" fontWeight="bold" sx={{ display: 'block', mb: 0.5 }}>
            REQUIRED FOR NODE BACKEND: SHARE SHEET WITH SERVICE ACCOUNT
          </Typography>
          <Typography variant="body2" color="text.secondary" sx={{ mb: 1 }}>
            If using Node server backend, open your Google Sheet in browser, click <strong>Share</strong>, and add this Service Account email with <strong>Editor</strong> access:
          </Typography>
          <Box sx={{ display: 'flex', alignItems: 'center', gap: 1, flexWrap: 'wrap' }}>
            <Typography variant="body2" fontFamily="monospace" fontWeight="bold" sx={{ bgcolor: 'background.paper', px: 1.5, py: 0.8, borderRadius: 1, border: '1px solid', borderColor: 'divider', wordBreak: 'break-all', flex: 1 }}>
              {SERVICE_ACCOUNT_EMAIL}
            </Typography>
            <Button
              variant="outlined"
              size="small"
              startIcon={<ContentCopy />}
              onClick={handleCopyServiceAccount}
            >
              Copy Email
            </Button>
          </Box>
        </Box>

        {/* Node Backend API Server URL Card */}
        <Box sx={{ mb: 3, p: 2, bgcolor: '#f0f4f9', borderRadius: 1.5, border: '1px solid #d3e3fd' }}>
          <Typography variant="caption" color="primary.main" fontWeight="bold" sx={{ display: 'block', mb: 0.5 }}>
            NODE BACKEND SERVER URL (HTTPS)
          </Typography>
          <Typography variant="body2" color="text.secondary" sx={{ mb: 1.5 }}>
            GitHub Pages is hosted over <strong>HTTPS</strong> and requires an <strong>HTTPS Server URL</strong> (e.g. cloud hosted Node server on Render/Railway or localtunnel) to allow other devices to connect to your Node backend.
          </Typography>
          
          <TextField
            fullWidth
            size="small"
            label="Node Server API URL (HTTPS)"
            placeholder="https://your-node-backend.onrender.com or https://funny-cat-42.loca.lt"
            value={apiUrlInput}
            onChange={(e) => {
              setApiUrlInput(e.target.value);
              setApiUrl(e.target.value);
            }}
            helperText="Paste your HTTPS Node server URL here so other devices can connect."
          />
        </Box>

        {/* Current Active Sheet Info */}
        <Box sx={{ mb: 3, p: 2, bgcolor: 'action.hover', borderRadius: 1.5 }}>
          <Typography variant="caption" color="text.secondary" fontWeight="bold" sx={{ display: 'block', mb: 0.5 }}>
            CURRENT ACTIVE DATABASE SHEET
          </Typography>
          <Box sx={{ display: 'flex', alignItems: 'center', gap: 1, flexWrap: 'wrap' }}>
            <Storage fontSize="small" color="action" />
            <Typography variant="body2" fontFamily="monospace" fontWeight="medium" sx={{ wordBreak: 'break-all', flex: 1 }}>
              {isLoadingConfig ? 'Loading...' : (activeSheetId || 'Not connected')}
            </Typography>
            {activeSheetUrl && (
              <Button
                variant="outlined"
                size="small"
                startIcon={<OpenInNew />}
                href={activeSheetUrl}
                target="_blank"
                rel="noopener noreferrer"
              >
                Open Sheet
              </Button>
            )}
          </Box>
        </Box>

        {/* Error Notice */}
        {errorMessage && (
          <Alert severity="error" sx={{ mb: 3 }} onClose={() => setErrorMessage(null)}>
            <Typography variant="subtitle2" fontWeight="bold">
              Database Connection Failed
            </Typography>
            <Typography variant="body2">
              {errorMessage}
            </Typography>
          </Alert>
        )}

        {/* Input Form */}
        <Box sx={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
          <Typography variant="subtitle2" fontWeight="bold">
            Paste New Google Sheet Link or ID
          </Typography>
          
          <TextField
            fullWidth
            placeholder="https://docs.google.com/spreadsheets/d/1kZh86P60Meu3YhKy7neFdlqfH6AJjuHSUfGqmvYvfVo/edit"
            value={sheetInput}
            onChange={(e) => setSheetInput(e.target.value)}
            disabled={isSavingSheet}
            InputProps={{
              startAdornment: (
                <InputAdornment position="start">
                  <LinkIcon color="action" />
                </InputAdornment>
              ),
              endAdornment: (
                <InputAdornment position="end">
                  <Tooltip title="Paste from Clipboard">
                    <IconButton onClick={handlePasteClipboard} edge="end" size="small">
                      <ContentPaste fontSize="small" />
                    </IconButton>
                  </Tooltip>
                </InputAdornment>
              )
            }}
            helperText="Paste the full URL from your browser address bar or just the Spreadsheet ID."
          />

          <Box sx={{ display: 'flex', gap: 2, justifyContent: 'flex-end', flexWrap: 'wrap' }}>
            <Button
              variant="outlined"
              color="inherit"
              startIcon={isTestingConn ? <CircularProgress size={18} /> : <Refresh />}
              onClick={handleTestConnection}
              disabled={isTestingConn || isSavingSheet}
            >
              Test Connection
            </Button>

            <Button
              variant="contained"
              color="primary"
              startIcon={isSavingSheet ? <CircularProgress size={18} color="inherit" /> : <Save />}
              onClick={handleSaveSpreadsheet}
              disabled={isSavingSheet || !sheetInput.trim()}
            >
              {isSavingSheet ? 'Connecting & Verifying...' : 'Save & Connect Database'}
            </Button>
          </Box>
        </Box>

        {/* Guidance Alert */}
        <Alert severity="info" sx={{ mt: 3, '& .MuiAlert-message': { width: '100%' } }}>
          <Typography variant="subtitle2" fontWeight="bold" gutterBottom>
            How to use your own Google Sheet:
          </Typography>
          <Typography variant="body2" component="div">
            <ol style={{ paddingLeft: 20, margin: 0 }}>
              <li>Create or open a Google Sheet in Google Drive.</li>
              <li>Share the sheet with your Service Account email: <code>{SERVICE_ACCOUNT_EMAIL}</code> (Editor access).</li>
              <li>Copy the sheet URL from your browser's address bar and paste it above.</li>
              <li>Click <strong>Save & Connect Database</strong> — the system will automatically create required tabs (<code>Users</code>, <code>Reports</code>, <code>Account Codes</code>, <code>Collections</code>, <code>Signatories</code>, <code>RPT Collections</code>) and headers if they do not exist!</li>
            </ol>
          </Typography>
        </Alert>
      </Paper>

      {/* Preferences Section */}
      <Paper sx={{ maxWidth: '100%', mb: 3, borderRadius: 2 }}>
        <List subheader={<Typography variant="overline" sx={{ px: 2, pt: 2, display: 'block', fontWeight: 'bold' }}>Preferences</Typography>}>
          <ListItem>
            <ListItemIcon>
              <Notifications />
            </ListItemIcon>
            <ListItemText primary="Notifications" secondary="Enable email notifications for new reports" />
            <Switch defaultChecked />
          </ListItem>
          <Divider variant="inset" component="li" />
          
          <ListItem>
            <ListItemIcon>
              <DarkMode />
            </ListItemIcon>
            <ListItemText primary="Dark Mode" secondary="Toggle dark/light theme" />
            <Switch />
          </ListItem>
          <Divider variant="inset" component="li" />
          
          <ListItem>
            <ListItemIcon>
              <Language />
            </ListItemIcon>
            <ListItemText primary="Language" secondary="English (US)" />
          </ListItem>
        </List>
      </Paper>

      {/* Data & Security Section */}
      <Paper sx={{ maxWidth: '100%', mb: 3, borderRadius: 2 }}>
        <List subheader={<Typography variant="overline" sx={{ px: 2, pt: 2, display: 'block', fontWeight: 'bold' }}>Data & Security</Typography>}>
            <ListItemButton onClick={handleExportData}>
                <ListItemIcon><CloudDownload /></ListItemIcon>
                <ListItemText primary="Export Local Data Backup" secondary="Download a JSON backup of your local storage data" />
            </ListItemButton>
            <Divider variant="inset" component="li" />

            <ListItemButton onClick={() => setOpenClearDialog(true)}>
                <ListItemIcon><DeleteForever color="error" /></ListItemIcon>
                <ListItemText 
                    primary="Clear Local Cache" 
                    secondary="Remove locally stored report cache (Caution)" 
                    primaryTypographyProps={{ color: 'error.main' }}
                />
            </ListItemButton>
            <Divider variant="inset" component="li" />
            
            <ListItem>
                <ListItemIcon>
                <Security />
                </ListItemIcon>
                <ListItemText primary="Security" secondary="Change password and security settings" />
            </ListItem>
        </List>
      </Paper>

      <Box sx={{ textAlign: 'center', py: 4, color: 'text.secondary' }}>
        <Typography variant="body2" sx={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 1 }}>
            <Info fontSize="small" /> RCD System v1.0.0
        </Typography>
        <Typography variant="caption">
            © 2024 Municipality of San Vicente. All rights reserved.
        </Typography>
      </Box>

      <ConfirmDialog
        open={openClearDialog}
        onClose={() => setOpenClearDialog(false)}
        onConfirm={handleClearData}
        title="Clear Local Data?"
        message="This will remove all reports stored in your browser's local storage. This action cannot be undone."
        confirmText="Clear Data"
        severity="error"
      />

      <Notification
        open={notification.open}
        onClose={() => setNotification({ ...notification, open: false })}
        message={notification.message}
        severity={notification.severity}
      />
    </Box>
  );
};

