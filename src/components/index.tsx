import React, { useState, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { 
  autoMapColumns, 
  convertToTallyJSON
} from './Tally/vouchers';
import type { 
  ColumnMapping, 
  ExcelRow 
} from './Tally/vouchers';
import '../css/style.css';

interface TallyField {
  key: string;
  label: string;
  required: boolean;
}

const FinPort: React.FC = () => {
  const [excelData, setExcelData] = useState<ExcelRow[]>([]);
  const [excelHeaders, setExcelHeaders] = useState<string[]>([]);
  const [generatedJSON, setGeneratedJSON] = useState<any>(null);
  const [currentPage, setCurrentPage] = useState<number>(1);
  const [pageSize, setPageSize] = useState<number>(50);
  const [totalPages, setTotalPages] = useState<number>(1);
  const [fileName, setFileName] = useState<string>('-');
  const [fileSize, setFileSize] = useState<string>('-');
  const [activeTab, setActiveTab] = useState<'data' | 'json'>('data');
  const [statusText, setStatusText] = useState<string>('Ready');
  const [statusType, setStatusType] = useState<'success' | 'warning' | 'error'>('success');
  const [toastMessage, setToastMessage] = useState<string>('');
  const [toastType, setToastType] = useState<'success' | 'error'>('success');
  const [showToast, setShowToast] = useState<boolean>(false);
  const [columnMapping, setColumnMapping] = useState<ColumnMapping>({});
  const [voucherType, setVoucherType] = useState<string>('Sales');
  const [defaultGodown, setDefaultGodown] = useState<string>('Main Location');
  const [salesLedger, setSalesLedger] = useState<string>('Sales');

  const fileInputRef = useRef<HTMLInputElement>(null);

  const tallyFields: TallyField[] = [
    { key: 'date', label: 'Date', required: true },
    { key: 'vouchernumber', label: 'Voucher No', required: true },
    { key: 'partyname', label: 'Party Name', required: true },
    { key: 'partygstno', label: 'Party GST', required: false },
    { key: 'stockitemname', label: 'Stock Item', required: true },
    { key: 'quantity', label: 'Quantity', required: true },
    { key: 'unit', label: 'Unit', required: false },
    { key: 'rate', label: 'Rate', required: true },
    { key: 'amount', label: 'Amount', required: true },
    { key: 'godownname', label: 'Godown', required: false },
    { key: 'batchname', label: 'Batch', required: false },
    { key: 'narration', label: 'Narration', required: false }
  ];

  useEffect(() => {
    calculatePagination();
  }, [excelData, pageSize]);

  useEffect(() => {
    if (showToast) {
      const timer = setTimeout(() => setShowToast(false), 3000);
      return () => clearTimeout(timer);
    }
  }, [showToast]);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) readExcelFile(file);
  };

  const readExcelFile = (file: File) => {
    updateStatus('Loading file...', 'warning');
    const reader = new FileReader();

    reader.onload = (e: ProgressEvent<FileReader>) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        if (jsonData.length > 0) {
          const headers = jsonData[0].map(h => String(h || '').trim());
          const rows = jsonData.slice(1).filter(row => row.some(cell => cell !== null && cell !== ''));

          setExcelHeaders(headers);
          setExcelData(rows);
          setFileName(file.name);
          setFileSize(formatFileSize(file.size));
          setCurrentPage(1);
          
          // Use the imported autoMapColumns function
          const mapping = autoMapColumns(headers);
          setColumnMapping(mapping);
          
          updateStatus('File loaded successfully', 'success');
          displayToast('Excel loaded: ' + rows.length + ' rows', 'success');
        }
      } catch (error: any) {
        updateStatus('Error loading file', 'error');
        displayToast('Error: ' + error.message, 'error');
      }
    };

    reader.readAsArrayBuffer(file);
  };

  const formatFileSize = (bytes: number): string => {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
  };

  const calculatePagination = () => {
    const total = Math.ceil(excelData.length / pageSize) || 1;
    setTotalPages(total);
  };

  const handleMappingChange = (fieldKey: string, columnIndex: string) => {
    setColumnMapping(prev => ({
      ...prev,
      [fieldKey]: columnIndex === '' ? -1 : parseInt(columnIndex)
    }));
  };

  const handleConvertToTallyJSON = () => {
    updateStatus('Converting to JSON...', 'warning');

    // Use the imported convertToTallyJSON function
    const json = convertToTallyJSON(
      excelData,
      columnMapping,
      voucherType,
      defaultGodown,
      salesLedger
    );

    setGeneratedJSON(json);
    setActiveTab('json');
    updateStatus('Converted ' + excelData.length + ' vouchers', 'success');
    displayToast('Converted ' + excelData.length + ' rows to Tally JSON', 'success');
  };

  const copyJSON = () => {
    if (!generatedJSON) return;
    navigator.clipboard.writeText(JSON.stringify(generatedJSON, null, 2));
    displayToast('JSON copied to clipboard', 'success');
  };

  const downloadJSON = () => {
    if (!generatedJSON) return;
    const blob = new Blob([JSON.stringify(generatedJSON, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'tally_import_' + new Date().toISOString().slice(0, 10) + '.json';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    displayToast('JSON file downloaded', 'success');
  };

  const downloadTemplate = () => {
    const templateData = [
      ['Date', 'Voucher No', 'Party Name', 'Party GST', 'Stock Item', 'Quantity', 'Unit', 'Rate', 'Amount', 'Godown', 'Batch', 'Narration'],
      ['2025-01-15', 'INV-001', 'ABC Trading Co', '27AABCU9603R1ZM', 'Product A', 100, 'Nos', 50, 5000, 'Main Godown', 'Batch-001', 'Sale of goods'],
      ['2025-01-15', 'INV-002', 'XYZ Enterprises', '27AABCU9603R1ZN', 'Product B', 50, 'Kgs', 120, 6000, 'Warehouse 1', 'Batch-002', 'Sale of materials']
    ];
    const ws = XLSX.utils.aoa_to_sheet(templateData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Template');
    XLSX.writeFile(wb, 'tally_import_template.xlsx');
    displayToast('Template downloaded', 'success');
  };

  const updateStatus = (text: string, type: 'success' | 'warning' | 'error') => {
    setStatusText(text);
    setStatusType(type);
  };

  const displayToast = (message: string, type: 'success' | 'error') => {
    setToastMessage(message);
    setToastType(type);
    setShowToast(true);
  };

  const renderTable = () => {
    const start = (currentPage - 1) * pageSize;
    const end = Math.min(start + pageSize, excelData.length);
    const pageData = excelData.slice(start, end);

    return (
      <table id="dataTable">
        <thead>
          <tr>
            <th>#</th>
            {excelHeaders.map((header, idx) => (
              <th key={idx}>{header}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {pageData.map((row, rowIdx) => (
            <tr key={start + rowIdx}>
              <td>{start + rowIdx + 1}</td>
              {excelHeaders.map((_, colIdx) => (
                <td key={colIdx}>
                  {row[colIdx] !== undefined ? String(row[colIdx]) : ''}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    );
  };

  return (
    <div style={{ height: '100vh', display: 'flex', flexDirection: 'column' }}>
      {/* Toolbar */}
      <div className="toolbar">
        <div className="app-title">
          <img src="./Logo.png" alt="" style={{ width: 30 }} />
          <span>FinPort</span>
        </div>

        <input
          type="file"
          id="fileInput"
          ref={fileInputRef}
          accept=".xlsx,.xls,.csv"
          onChange={handleFileUpload}
          style={{ display: 'none' }}
        />
        <button className="toolbar-btn" onClick={() => fileInputRef.current?.click()}>
          <i className="ti ti-file-upload"></i>
          Upload Excel
        </button>

        <button className="toolbar-btn" onClick={downloadTemplate}>
          <i className="ti ti-file-download"></i>
          Template
        </button>

        <div className="toolbar-separator"></div>

        <button
          className="toolbar-btn primary"
          onClick={handleConvertToTallyJSON}
          disabled={excelData.length === 0}
        >
          <i className="ti ti-refresh"></i>
          Convert to JSON
        </button>

        <button
          className="toolbar-btn success"
          onClick={downloadJSON}
          style={{ display: generatedJSON ? 'inline-flex' : 'none' }}
        >
          <i className="ti ti-download"></i>
          Download JSON
        </button>

        <div className="toolbar-separator"></div>

        <span className="toolbar-label">Voucher:</span>
        <select
          className="toolbar-select"
          value={voucherType}
          onChange={(e) => setVoucherType(e.target.value)}
        >
          <option value="Sales">Sales</option>
          <option value="Sales Busy">Sales Busy</option>
          <option value="Purchase">Purchase</option>
          <option value="Purchase Busy">Purchase Busy</option>
          <option value="Receipt">Receipt</option>
          <option value="Payment">Payment</option>
          <option value="Journal">Journal</option>
          <option value="Contra">Contra</option>
          <option value="Credit Note">Credit Note</option>
          <option value="Debit Note">Debit Note</option>
        </select>

        <span className="toolbar-label">Godown:</span>
        <input
          type="text"
          className="toolbar-input"
          value={defaultGodown}
          onChange={(e) => setDefaultGodown(e.target.value)}
          placeholder="Godown Name"
        />

        <span className="toolbar-label">Sales Ledger:</span>
        <input
          type="text"
          className="toolbar-input"
          value={salesLedger}
          onChange={(e) => setSalesLedger(e.target.value)}
          placeholder="Sales Ledger"
        />
      </div>

      {/* Main Content */}
      <div className="main-container">
        {/* Tab Bar */}
        <div className="tab-bar">
          <div
            className={`tab-item ${activeTab === 'data' ? 'active' : ''}`}
            onClick={() => setActiveTab('data')}
          >
            <i className="ti ti-table"></i>
            Data Preview
          </div>
          <div
            className={`tab-item ${activeTab === 'json' ? 'active' : ''}`}
            onClick={() => setActiveTab('json')}
          >
            <i className="ti ti-code"></i>
            JSON Output
          </div>
        </div>

        {/* Info Bar */}
        <div className={`info-bar ${excelData.length > 0 ? 'show' : ''}`}>
          <div className="info-item">
            <i className="ti ti-file-spreadsheet"></i>
            <label>File:</label>
            <span>{fileName}</span>
          </div>
          <div className="info-item">
            <i className="ti ti-list-numbers"></i>
            <label>Rows:</label>
            <span>{excelData.length.toLocaleString()}</span>
          </div>
          <div className="info-item">
            <i className="ti ti-columns"></i>
            <label>Columns:</label>
            <span>{excelHeaders.length}</span>
          </div>
          <div className="info-item">
            <i className="ti ti-database"></i>
            <label>Size:</label>
            <span>{fileSize}</span>
          </div>
        </div>

        {/* Mapping Panel */}
        <div className={`mapping-panel ${excelData.length > 0 ? 'show' : ''}`}>
          <div className="mapping-title">
            <i className="ti ti-link"></i>
            Column Mapping - Map Excel columns to Tally fields
          </div>
          <div className="mapping-grid">
            {tallyFields.map((field) => (
              <div className="mapping-item" key={field.key}>
                <label className={field.required ? 'required' : ''}>
                  {field.label}
                </label>
                <select
                  value={columnMapping[field.key] ?? ''}
                  onChange={(e) => handleMappingChange(field.key, e.target.value)}
                >
                  <option value="">-- Select --</option>
                  {excelHeaders.map((header, index) => (
                    <option key={index} value={index}>
                      {header}
                    </option>
                  ))}
                </select>
              </div>
            ))}
          </div>
        </div>

        {/* Data Container */}
        <div className={`data-container ${activeTab === 'json' ? 'hidden' : ''}`}>
          <div className="table-wrapper">
            <div className="table-container">
              {excelData.length === 0 ? (
                <div className="empty-state">
                  <i className="ti ti-file-spreadsheet"></i>
                  <h3>No Data Loaded</h3>
                  <p>Click "Upload Excel" to import your file</p>
                </div>
              ) : (
                renderTable()
              )}
            </div>
          </div>

          {/* Pagination */}
          {excelData.length > 0 && (
            <div className="pagination-bar">
              <div className="pagination-info">
                Showing <strong>{(currentPage - 1) * pageSize + 1}</strong> to{' '}
                <strong>{Math.min(currentPage * pageSize, excelData.length)}</strong> of{' '}
                <strong>{excelData.length.toLocaleString()}</strong> records
              </div>
              <div className="pagination-controls">
                <span>Rows:</span>
                <select
                  className="page-select"
                  value={pageSize}
                  onChange={(e) => {
                    setPageSize(parseInt(e.target.value));
                    setCurrentPage(1);
                  }}
                >
                  <option value="25">25</option>
                  <option value="50">50</option>
                  <option value="100">100</option>
                  <option value="200">200</option>
                  <option value="500">500</option>
                </select>
                <div className="toolbar-separator" style={{ height: 18, margin: '0 10px' }}></div>
                <button
                  className="page-btn"
                  onClick={() => setCurrentPage(1)}
                  disabled={currentPage === 1}
                >
                  <i className="ti ti-chevrons-left"></i>
                </button>
                <button
                  className="page-btn"
                  onClick={() => setCurrentPage(prev => Math.max(1, prev - 1))}
                  disabled={currentPage === 1}
                >
                  <i className="ti ti-chevron-left"></i>
                </button>
                <span>
                  Page{' '}
                  <input
                    type="number"
                    className="page-input"
                    value={currentPage}
                    min="1"
                    max={totalPages}
                    onChange={(e) => {
                      const page = parseInt(e.target.value);
                      if (page >= 1 && page <= totalPages) {
                        setCurrentPage(page);
                      }
                    }}
                  />{' '}
                  of <strong>{totalPages}</strong>
                </span>
                <button
                  className="page-btn"
                  onClick={() => setCurrentPage(prev => Math.min(totalPages, prev + 1))}
                  disabled={currentPage === totalPages}
                >
                  <i className="ti ti-chevron-right"></i>
                </button>
                <button
                  className="page-btn"
                  onClick={() => setCurrentPage(totalPages)}
                  disabled={currentPage === totalPages}
                >
                  <i className="ti ti-chevrons-right"></i>
                </button>
              </div>
            </div>
          )}
        </div>

        {/* JSON Container */}
        <div className={`json-container ${activeTab === 'json' ? 'active' : ''}`}>
          <div className="json-toolbar">
            <button className="toolbar-btn" onClick={copyJSON}>
              <i className="ti ti-copy"></i>
              Copy to Clipboard
            </button>
            <button className="toolbar-btn success" onClick={downloadJSON}>
              <i className="ti ti-download"></i>
              Download JSON
            </button>
          </div>
          <div className="json-content">
            {generatedJSON
              ? JSON.stringify(generatedJSON, null, 2)
              : '// JSON output will appear here after conversion'}
          </div>
        </div>
      </div>

      {/* Status Bar */}
      <div className="status-bar">
        <div className="status-left">
          <div className="status-item">
            <span className={`status-indicator ${statusType}`}></span>
            <span>{statusText}</span>
          </div>
        </div>
        <div className="status-right">
          <span>Excel to Tally JSON Converter v1.0</span>
        </div>
      </div>

      {/* Toast Notification */}
      <div className={`toast ${toastType} ${showToast ? 'show' : ''}`}>
        <i className={toastType === 'success' ? 'ti ti-check' : 'ti ti-alert-circle'}></i>
        <span>{toastMessage}</span>
      </div>
    </div>
  );
};

export default FinPort;