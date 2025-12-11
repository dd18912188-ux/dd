import React, { useState, useMemo } from 'react';
import { createRoot } from 'react-dom/client';

// Define global XLSX variable provided by the CDN
declare const XLSX: any;

// --- Inline Icons (Replaces Lucide dependency for better performance) ---
const Icons = {
    FileSpreadsheet: ({ className }: { className?: string }) => (
        <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}><path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z"/><path d="M14 2v4a2 2 0 0 0 2 2h4"/><path d="M8 13h2"/><path d="M14 13h2"/><path d="M8 17h2"/><path d="M14 17h2"/></svg>
    ),
    UploadCloud: ({ className }: { className?: string }) => (
        <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}><path d="M4 14.899A7 7 0 1 1 15.71 8h1.79a4.5 4.5 0 0 1 2.5 8.242"/><path d="M12 12v9"/><path d="m16 16-4-4-4 4"/></svg>
    ),
    CheckCircle: ({ className }: { className?: string }) => (
        <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>
    ),
    Loader2: ({ className }: { className?: string }) => (
        <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}><path d="M21 12a9 9 0 1 1-6.219-8.56"/></svg>
    ),
    BarChart2: ({ className }: { className?: string }) => (
        <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}><line x1="18" x2="18" y1="20" y2="10"/><line x1="12" x2="12" y1="20" y2="4"/><line x1="6" x2="6" y1="20" y2="14"/></svg>
    ),
    AlertCircle: ({ className }: { className?: string }) => (
        <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}><circle cx="12" cy="12" r="10"/><line x1="12" x2="12" y1="8" y2="12"/><line x1="12" x2="12.01" y1="16" y2="16"/></svg>
    ),
    Calendar: ({ className }: { className?: string }) => (
        <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}><rect width="18" height="18" x="3" y="4" rx="2" ry="2"/><line x1="16" x2="16" y1="2" y2="6"/><line x1="8" x2="8" y1="2" y2="6"/><line x1="3" x2="21" y1="10" y2="10"/></svg>
    ),
    User: ({ className }: { className?: string }) => (
        <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}><path d="M19 21v-2a4 4 0 0 0-4-4H9a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/></svg>
    ),
    Ruler: ({ className }: { className?: string }) => (
        <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}><path d="M21.3 15.3a2.4 2.4 0 0 1 0 3.4l-2.6 2.6a2.4 2.4 0 0 1-3.4 0L2.7 8.7a2.41 2.41 0 0 1 0-3.4l2.6-2.6a2.41 2.41 0 0 1 3.4 0Z"/><path d="m14.5 12.5 2-2"/><path d="m11.5 9.5 2-2"/><path d="m8.5 6.5 2-2"/><path d="m17.5 15.5 2-2"/></svg>
    ),
    Calculator: ({ className }: { className?: string }) => (
        <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}><rect width="16" height="20" x="4" y="2" rx="2"/><line x1="8" x2="16" y1="6" y2="6"/><line x1="16" x2="16" y1="14" y2="18"/><path d="M16 10h.01"/><path d="M12 10h.01"/><path d="M8 10h.01"/><path d="M12 14h.01"/><path d="M8 14h.01"/><path d="M12 18h.01"/><path d="M8 18h.01"/></svg>
    ),
    Table: ({ className }: { className?: string }) => (
        <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}><path d="M12 3v18"/><rect width="18" height="18" x="3" y="3" rx="2"/><path d="M3 9h18"/><path d="M3 15h18"/></svg>
    ),
    Download: ({ className }: { className?: string }) => (
        <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" x2="12" y1="15" y2="3"/></svg>
    ),
    RefreshCw: ({ className }: { className?: string }) => (
         <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}><path d="M3 12a9 9 0 0 1 9-9 9.75 9.75 0 0 1 6.74 2.74L21 8"/><path d="M21 3v5h-5"/><path d="M21 12a9 9 0 0 1-9 9 9.75 9.75 0 0 1-6.74-2.74L3 16"/><path d="M8 16H3v5"/></svg>
    )
};

// --- Main App Component ---
const App = () => {
    const [file, setFile] = useState<File | null>(null);
    const [processedData, setProcessedData] = useState<any[]>([]);
    const [isProcessing, setIsProcessing] = useState(false);
    const [errorMsg, setErrorMsg] = useState("");
    const [fileName, setFileName] = useState("");
    const [metadata, setMetadata] = useState({ unit: "", globalDate: "", globalCustomer: "", batchSource: "" });
    const [dragActive, setDragActive] = useState(false);

    // Field mapping configuration
    const fieldMappings: Record<string, string[]> = useMemo(() => ({
        date: ['日期', 'Date', '时间', '交期'], 
        customer: ['客户', '客户名称', 'Client', 'Customer'],
        contract: ['合同', '合同编号', '订单', 'PO'],
        product: ['产品', '产品编号', '款号', '货号', '品名'],
        spec: ['门幅', '克重', '规格', '幅宽'],
        color: ['颜色', '色号', 'Color'],
        batch: ['缸号', '批号', 'Batch'], 
        quantity: ['数量', 'Qty', '数'], 
        weight: ['重量', 'KG', '公斤', '净重'],
        roll: ['匹号', '卷号', '件号', 'Roll'],
        remarks: ['备注', '说明', 'Note']
    }), []);

    // --- Handlers ---
    const handleDrag = (e: React.DragEvent) => {
        e.preventDefault();
        e.stopPropagation();
        if (e.type === "dragenter" || e.type === "dragover") {
            setDragActive(true);
        } else if (e.type === "dragleave") {
            setDragActive(false);
        }
    };

    const handleDrop = (e: React.DragEvent) => {
        e.preventDefault();
        e.stopPropagation();
        setDragActive(false);
        if (e.dataTransfer.files && e.dataTransfer.files[0]) {
            processFile(e.dataTransfer.files[0]);
        }
    };

    const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
        if (e.target.files && e.target.files[0]) {
            processFile(e.target.files[0]);
        }
    };

    const processFile = (selectedFile: File) => {
        if (!selectedFile.name.match(/\.(xlsx|xls)$/)) {
            setErrorMsg("请上传 Excel 文件 (.xlsx 或 .xls)");
            return;
        }
        setFile(selectedFile);
        setFileName(selectedFile.name);
        setErrorMsg("");
        setProcessedData([]);
        setMetadata({ unit: "", globalDate: "", globalCustomer: "", batchSource: "" });
    };

    const resetAll = () => {
        setFile(null);
        setFileName("");
        setProcessedData([]);
        setErrorMsg("");
        setMetadata({ unit: "", globalDate: "", globalCustomer: "", batchSource: "" });
    };

    // --- Business Logic ---
    const processExcel = () => {
        if (!file) return;
        
        setIsProcessing(true);
        const reader = new FileReader();

        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target?.result as ArrayBuffer);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                
                // Helper to get raw cell value
                const getCellValue = (address: string) => {
                    const cell = worksheet[address];
                    return cell ? (cell.v !== undefined ? cell.v : cell.w) : null;
                };

                // Legacy requirement: Check C9 for date
                const rawC9 = getCellValue('C9');
                
                // Convert sheet to JSON array of arrays (header: 1 means array of arrays)
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                
                if (jsonData.length < 1) throw new Error("文件内容为空");

                // Simulate AI processing delay for UX (feels more "calculated")
                setTimeout(() => {
                    analyzeAndTransform(jsonData, rawC9);
                    setIsProcessing(false);
                }, 800);

            } catch (err) {
                console.error(err);
                setErrorMsg("读取文件失败，请确保文件未加密且格式正确。");
                setIsProcessing(false);
            }
        };
        reader.readAsArrayBuffer(file);
    };

    const formatExcelDate = (val: any) => {
        if (!val) return "";
        // Check if it's an Excel serial date
        if (typeof val === 'number' && val > 20000) {
            const date = XLSX.SSF.parse_date_code(val);
            if (date) return `${date.y}-${String(date.m).padStart(2, '0')}-${String(date.d).padStart(2, '0')}`;
        }
        // Try parsing string date
        const strVal = String(val).trim();
        // Simple regex for YYYY/MM/DD or YYYY-MM-DD to standard format
        if (strVal.match(/^\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2}$/)) {
            return strVal.replace(/\//g, '-');
        }
        return strVal;
    };

    const extractHeaderMetadata = (rows: any[], headerRowIndex: number) => {
        let globalDate = "";
        let globalCustomer = "";
        
        // Scan rows above header for metadata
        for (let i = 0; i < headerRowIndex; i++) {
            const row = rows[i];
            if (!row || row.length === 0) continue;
            
            row.forEach((cell: any, cellIdx: number) => {
                if (!cell) return;
                const text = String(cell).trim().replace(/[:：]/g, ":");
                
                if (text.includes("日期") || text.includes("Date")) {
                    const parts = text.split(":");
                    if (parts.length > 1 && parts[1].trim()) {
                        globalDate = formatExcelDate(parts[1].trim());
                    } else if (row[cellIdx + 1]) {
                        globalDate = formatExcelDate(row[cellIdx + 1]);
                    }
                }
                
                if (text.includes("客户") || text.includes("Client")) {
                    const parts = text.split(":");
                    if (parts.length > 1 && parts[1].trim()) {
                        globalCustomer = parts[1].trim();
                    } else if (row[cellIdx + 1]) {
                        globalCustomer = String(row[cellIdx + 1]).trim();
                    }
                }
            });
        }
        return { globalDate, globalCustomer };
    };

    const analyzeAndTransform = (rows: any[], rawC9: any) => {
        let headerRowIndex = -1;
        let columnMap: Record<string, number> = {}; 
        let extractedUnit = "";
        const fixedDate = rawC9 ? formatExcelDate(rawC9) : "";

        // 1. Identify Header Row
        for (let i = 0; i < Math.min(rows.length, 20); i++) {
            const row = rows[i];
            let matchCount = 0;
            if (!row) continue;
            
            row.forEach((cell: any) => {
                if (typeof cell !== 'string') return;
                const cellStr = cell.trim();
                
                // Check against all field mappings
                Object.keys(fieldMappings).forEach(key => {
                    if (fieldMappings[key].some(keyword => cellStr.includes(keyword))) matchCount++;
                });

                // Attempt to detect unit from Quantity column header, e.g., "数量(米)"
                if (cellStr.includes("数量") || cellStr.toLowerCase().includes("qty")) {
                    if (cellStr.includes("米")) extractedUnit = "米";
                    else if (cellStr.includes("码")) extractedUnit = "码";
                    else if (cellStr.includes("KG") || cellStr.includes("kg")) extractedUnit = "KG";
                    else {
                        const match = cellStr.match(/[\(（](.*?)[\)）]/);
                        if (match && match[1]) extractedUnit = match[1];
                    }
                }
            });

            // If enough keywords match, assume this is the header
            if (matchCount >= 2) { 
                headerRowIndex = i;
                // Map columns
                row.forEach((cell: any, cellIndex: number) => {
                    if (!cell) return;
                    const cellStr = cell.toString().trim();
                    Object.keys(fieldMappings).forEach(key => {
                        if (fieldMappings[key].some(keyword => cellStr.includes(keyword))) {
                            // Only map if not already mapped (prefer first occurrence)
                            if (columnMap[key] === undefined) columnMap[key] = cellIndex;
                        }
                    });
                });
                break; 
            }
        }

        if (headerRowIndex === -1) {
            setErrorMsg("AI无法自动识别表头，请确保表格包含标准的列名（如：缸号、数量、客户等）。");
            return;
        }

        // 2. Determine Batch Column Source
        let batchColIndex = columnMap['batch'];
        let sourceInfo = "自动识别";
        if (batchColIndex === undefined) {
            // Fallback to Column F (index 5) if standard keywords aren't found
            batchColIndex = 5;
            sourceInfo = "默认 F 列 (未找到表头)";
        }

        // 3. Extract Metadata
        const scannedMetadata = extractHeaderMetadata(rows, headerRowIndex);
        const finalDate = fixedDate || scannedMetadata.globalDate;
        const finalCustomer = scannedMetadata.globalCustomer; 
        const finalUnit = extractedUnit || "米";

        setMetadata({ unit: finalUnit, globalDate: finalDate, globalCustomer: finalCustomer, batchSource: sourceInfo });

        // 4. Process Data Rows
        const rawDataRows = rows.slice(headerRowIndex + 1);
        const batchGroups: Record<string, any> = {};
        let lastValidBatch: string | null = null; 

        rawDataRows.forEach(row => {
            const rowString = row.map((cell: any) => String(cell || "")).join("").toLowerCase();
            // Skip footer/signature rows
            if (rowString.includes("签名") || rowString.includes("signature") || rowString.includes("盖章") || rowString.includes("核对") || rowString.includes("approved")) return;

            const qtyIdx = columnMap['quantity'];
            const rollIdx = columnMap['roll']; 
            const weightIdx = columnMap['weight'];
            
            const rawQty = qtyIdx !== undefined ? row[qtyIdx] : null;
            const rawRoll = rollIdx !== undefined ? row[rollIdx] : null;
            
            // Skip rows without quantity or roll number
            const hasData = (rawQty !== undefined && rawQty !== null) || (rawRoll !== undefined && rawRoll !== null);
            if (!hasData) return;

            // Handle Batch Merging (Empty batch cell inherits from above)
            let rawBatch = row[batchColIndex];
            let currentBatchStr = "";

            if (rawBatch !== undefined && rawBatch !== null && String(rawBatch).trim() !== "") {
                currentBatchStr = String(rawBatch).trim();
                // Stop if we hit a subtotal row inside the data
                if (currentBatchStr.includes("计") || currentBatchStr.includes("Subtotal") || currentBatchStr.includes("Total")) {
                    lastValidBatch = null; 
                    return; 
                }
                lastValidBatch = currentBatchStr;
            } else if (lastValidBatch) {
                currentBatchStr = lastValidBatch;
            }

            if (!currentBatchStr) currentBatchStr = "未知缸号";
            const key = currentBatchStr;
           
            if (!batchGroups[key]) {
                batchGroups[key] = {
                    date: finalDate || formatExcelDate(getValue(row, columnMap['date'])),
                    customer: finalCustomer || getValue(row, columnMap['customer']),
                    contract: getValue(row, columnMap['contract']),
                    product: getValue(row, columnMap['product']),
                    spec: getValue(row, columnMap['spec']),
                    color: getValue(row, columnMap['color']),
                    batch: key,
                    quantity: 0,
                    unit: finalUnit,
                    weight: 0,
                    remarks: getValue(row, columnMap['remarks']),
                };
            }

            // Accumulate numeric values
            const qty = parseFloat(getValue(row, qtyIdx));
            if (!isNaN(qty)) batchGroups[key].quantity += qty;
            
            const w = parseFloat(getValue(row, weightIdx));
            if (!isNaN(w)) batchGroups[key].weight += w;
        });

        // 5. Finalize Result
        const finalResult = Object.values(batchGroups).map(item => ({
            ...item,
            quantity: parseFloat(item.quantity.toFixed(2)),
            weight: parseFloat(item.weight.toFixed(2)),
        }));

        setProcessedData(finalResult);
    };

    const getValue = (row: any[], index: number) => {
        if (index === undefined || row[index] === undefined) return "";
        return row[index];
    };

    const exportToExcel = () => {
        if (processedData.length === 0) return;
        const header = ["日期", "客户名称", "合同编号", "产品编号", "门幅/克重", "颜色", "缸号", "数量(总和)", "单位", "重量(KG)", "备注"];
        const dataRows = processedData.map(d => [d.date, d.customer, d.contract, d.product, d.spec, d.color, d.batch, d.quantity, d.unit, d.weight, d.remarks]);
        const ws = XLSX.utils.aoa_to_sheet([header, ...dataRows]);
        
        // Set column widths
        ws['!cols'] = [{wch: 12}, {wch: 20}, {wch: 15}, {wch: 15}, {wch: 15}, {wch: 10}, {wch: 15}, {wch: 12}, {wch: 8}, {wch: 10}, {wch: 20}];
        
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "统计结果");
        XLSX.writeFile(wb, `数据统计结果_${new Date().getTime()}.xlsx`);
    };

    const totalQuantity = processedData.reduce((acc, curr) => acc + (curr.quantity || 0), 0);
    const totalWeight = processedData.reduce((acc, curr) => acc + (curr.weight || 0), 0);

    // --- Render ---
    return (
        <div className="min-h-screen p-4 md:p-8 max-w-7xl mx-auto font-sans">
            {/* Header */}
            <div className="mb-8 text-center animate-fade-in">
                <h1 className="text-3xl md:text-4xl font-extrabold text-slate-800 mb-3 flex items-center justify-center gap-3">
                    <span className="p-2 bg-blue-600 rounded-lg text-white shadow-lg">
                        <Icons.FileSpreadsheet className="w-8 h-8" />
                    </span>
                    AI 智能单据统计
                </h1>
                <p className="text-slate-500 text-sm md:text-base max-w-2xl mx-auto">
                    自动识别并汇总 Excel 单据中的缸号、数量和重量。支持自动提取表头日期的 C9 规则。
                </p>
            </div>

            {/* Main Input Area */}
            <div className="glass-panel rounded-2xl p-6 md:p-8 shadow-xl mb-8 animate-fade-in" style={{animationDelay: '0.1s'}}>
                <div className="flex flex-col md:flex-row items-stretch justify-center gap-6">
                    
                    {/* File Drop Zone */}
                    <div 
                        className={`w-full md:w-1/2 relative group min-h-[160px] md:min-h-[200px]`}
                        onDragEnter={handleDrag} 
                        onDragLeave={handleDrag} 
                        onDragOver={handleDrag} 
                        onDrop={handleDrop}
                    >
                        <input 
                            type="file" 
                            accept=".xlsx, .xls" 
                            onChange={handleFileChange} 
                            className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10" 
                            disabled={isProcessing}
                        />
                        <div className={`drop-zone absolute inset-0 rounded-2xl border-2 border-dashed flex flex-col items-center justify-center p-6 text-center transition-all duration-300 ${dragActive ? 'active border-blue-500 bg-blue-50' : (fileName ? 'border-emerald-500 bg-emerald-50/50' : 'border-slate-300 bg-white/50 hover:border-blue-400')}`}>
                            {fileName ? (
                                <div className="flex flex-col items-center justify-center text-emerald-700 animate-fade-in">
                                    <Icons.CheckCircle className="w-12 h-12 mb-2" />
                                    <span className="font-bold text-lg truncate max-w-[250px]">{fileName}</span>
                                    <span className="text-xs text-emerald-600 mt-1">点击更换文件</span>
                                </div>
                            ) : (
                                <div className="text-slate-400 pointer-events-none">
                                    <div className="bg-white p-3 rounded-full shadow-sm inline-block mb-3">
                                        <Icons.UploadCloud className="w-8 h-8 text-blue-500" />
                                    </div>
                                    <p className="font-medium text-slate-600">点击或拖拽上传 Excel</p>
                                    <p className="text-xs mt-1 text-slate-400">支持 .xlsx, .xls 格式</p>
                                </div>
                            )}
                        </div>
                    </div>

                    {/* Actions Area */}
                    <div className="w-full md:w-1/2 flex flex-col justify-center gap-4">
                        <button 
                            onClick={processExcel} 
                            disabled={isProcessing || !file} 
                            className={`w-full py-4 rounded-xl font-bold text-white shadow-lg flex items-center justify-center gap-2 transition-all transform hover:translate-y-[-2px] active:scale-95 ${
                                isProcessing 
                                ? 'bg-slate-400 cursor-not-allowed' 
                                : (file ? 'bg-gradient-to-r from-blue-600 to-indigo-600 hover:shadow-blue-500/30' : 'bg-slate-300 cursor-not-allowed')
                            }`}
                        >
                            {isProcessing ? (
                                <><Icons.Loader2 className="w-5 h-5 animate-spin" /> AI 分析中...</>
                            ) : (
                                <><Icons.BarChart2 className="w-5 h-5" /> 开始智能统计</>
                            )}
                        </button>
                        
                        {fileName && (
                            <button 
                                onClick={resetAll} 
                                className="w-full py-3 rounded-xl border border-slate-200 text-slate-600 hover:bg-slate-50 hover:text-slate-800 transition-colors flex items-center justify-center gap-2 text-sm font-medium"
                            >
                                <Icons.RefreshCw className="w-4 h-4" /> 重置
                            </button>
                        )}
                    </div>
                </div>

                {/* Error Message */}
                {errorMsg && (
                    <div className="mt-6 p-4 bg-red-50 border border-red-100 text-red-700 rounded-xl text-center flex items-center justify-center gap-2 animate-fade-in">
                        <Icons.AlertCircle className="w-5 h-5 flex-shrink-0" />
                        <span>{errorMsg}</span>
                    </div>
                )}
                
                {/* Metadata Cards */}
                {!isProcessing && (metadata.unit || metadata.globalDate || metadata.globalCustomer || processedData.length > 0) && (
                    <div className="mt-8 grid grid-cols-2 md:grid-cols-4 gap-3 animate-fade-in">
                        <div className="bg-blue-50/80 p-3 rounded-xl border border-blue-100 flex flex-col items-center justify-center text-center">
                            <div className="flex items-center gap-1 text-blue-600 text-xs font-bold uppercase mb-1"><Icons.Calendar className="w-3 h-3" />日期</div>
                            <div className="text-blue-900 font-bold truncate w-full">{metadata.globalDate || "-"}</div>
                        </div>
                        <div className="bg-purple-50/80 p-3 rounded-xl border border-purple-100 flex flex-col items-center justify-center text-center">
                            <div className="flex items-center gap-1 text-purple-600 text-xs font-bold uppercase mb-1"><Icons.User className="w-3 h-3" />客户</div>
                            <div className="text-purple-900 font-bold truncate w-full">{metadata.globalCustomer || "-"}</div>
                        </div>
                        <div className="bg-emerald-50/80 p-3 rounded-xl border border-emerald-100 flex flex-col items-center justify-center text-center">
                            <div className="flex items-center gap-1 text-emerald-600 text-xs font-bold uppercase mb-1"><Icons.Ruler className="w-3 h-3" />单位</div>
                            <div className="text-emerald-900 font-bold truncate w-full">{metadata.unit || "米"}</div>
                        </div>
                        <div className="bg-orange-50/80 p-3 rounded-xl border border-orange-100 flex flex-col items-center justify-center text-center">
                            <div className="flex items-center gap-1 text-orange-600 text-xs font-bold uppercase mb-1"><Icons.Calculator className="w-3 h-3" />总数</div>
                            <div className="text-orange-900 font-bold truncate w-full">{totalQuantity.toLocaleString(undefined, {minimumFractionDigits: 0, maximumFractionDigits: 2})}</div>
                        </div>
                    </div>
                )}
            </div>

            {/* Results Table */}
            {processedData.length > 0 && (
                <div className="bg-white rounded-2xl shadow-xl overflow-hidden animate-fade-in border border-slate-100" style={{animationDelay: '0.2s'}}>
                    <div className="p-5 border-b border-slate-100 flex flex-col sm:flex-row justify-between items-center bg-slate-50/50 gap-4">
                        <div className="flex items-center gap-2">
                            <span className="bg-white p-2 rounded-lg shadow-sm text-slate-700"><Icons.Table className="w-5 h-5" /></span>
                            <div>
                                <h2 className="font-bold text-slate-800 text-lg">统计结果</h2>
                                <p className="text-xs text-slate-500">共合并 {processedData.length} 个缸号 (总重: {totalWeight.toFixed(2)} KG)</p>
                            </div>
                        </div>
                        <button 
                            onClick={exportToExcel} 
                            className="px-5 py-2.5 bg-emerald-600 text-white text-sm font-semibold rounded-xl hover:bg-emerald-700 hover:shadow-lg hover:shadow-emerald-200 transition-all flex items-center gap-2"
                        >
                            <Icons.Download className="w-4 h-4" /> 导出 Excel
                        </button>
                    </div>
                    
                    <div className="overflow-x-auto custom-scrollbar max-h-[600px]">
                        <table className="w-full text-sm text-left text-slate-600">
                            <thead className="text-xs text-slate-500 uppercase bg-slate-50 sticky top-0 z-10 shadow-sm">
                                <tr>
                                    <th className="px-6 py-4 font-semibold">日期</th>
                                    <th className="px-6 py-4 font-semibold">客户名称</th>
                                    <th className="px-6 py-4 font-semibold">合同/订单</th>
                                    <th className="px-6 py-4 font-semibold">产品编号</th>
                                    <th className="px-6 py-4 font-semibold">规格</th>
                                    <th className="px-6 py-4 font-semibold">颜色</th>
                                    <th className="px-6 py-4 font-semibold bg-amber-50 text-amber-800 border-b-2 border-amber-100">缸号</th>
                                    <th className="px-6 py-4 font-semibold text-right bg-blue-50 text-blue-800 border-b-2 border-blue-100">数量 (总和)</th>
                                    <th className="px-6 py-4 font-semibold text-center">单位</th>
                                    <th className="px-6 py-4 font-semibold text-right">重量 (KG)</th>
                                    <th className="px-6 py-4 font-semibold">备注</th>
                                </tr>
                            </thead>
                            <tbody className="divide-y divide-slate-100">
                                {processedData.map((row, idx) => (
                                    <tr key={idx} className="hover:bg-slate-50 transition-colors group">
                                        <td className="px-6 py-3 whitespace-nowrap text-slate-500">{row.date}</td>
                                        <td className="px-6 py-3 whitespace-nowrap font-medium text-slate-900">{row.customer}</td>
                                        <td className="px-6 py-3 whitespace-nowrap">{row.contract}</td>
                                        <td className="px-6 py-3 whitespace-nowrap">{row.product}</td>
                                        <td className="px-6 py-3 whitespace-nowrap">{row.spec}</td>
                                        <td className="px-6 py-3 whitespace-nowrap">
                                            <span className="inline-block px-2 py-1 rounded bg-slate-100 text-slate-600 text-xs">{row.color}</span>
                                        </td>
                                        <td className="px-6 py-3 whitespace-nowrap font-bold text-amber-900 bg-amber-50/30 group-hover:bg-amber-50/60">{row.batch}</td>
                                        <td className="px-6 py-3 text-right font-bold text-blue-600 bg-blue-50/30 group-hover:bg-blue-50/60">{row.quantity.toLocaleString()}</td>
                                        <td className="px-6 py-3 text-center text-xs text-slate-400">{row.unit}</td>
                                        <td className="px-6 py-3 text-right font-mono text-slate-700">{row.weight > 0 ? row.weight.toFixed(2) : '-'}</td>
                                        <td className="px-6 py-3 text-slate-400 text-xs max-w-[200px] truncate" title={row.remarks}>{row.remarks}</td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>
                </div>
            )}
        </div>
    );
};

const root = createRoot(document.getElementById('root')!);
root.render(<App />);
