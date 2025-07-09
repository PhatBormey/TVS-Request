
import React, { useState, useCallback, useEffect, useMemo, useRef } from 'react';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import { GoogleGenAI } from "@google/genai";
import * as pdfjsLib from 'pdfjs-dist';
import type { Report } from './types';
import { STATIONARY_ITEMS_ROW1, STATIONARY_ITEMS_ROW2, CAMPUS_OPTIONS } from './constants';

interface StockItem {
    quantity: number;
    lastInDate: string;
    lastOutDate: string;
    lastUpdateQuantity: number;
}

const initialFormData: Omit<Report, 'id'> = {
  requesterName: '',
  campus: '',
  importDate: '',
  exportDate: '',
  items: {},
  status: 'Process',
};

const initialReports: Report[] = [];

const LOCAL_STORAGE_KEY_REPORTS = 'stationaryAppReports';
const LOCAL_STORAGE_KEY_FORM_DATA = 'stationaryAppFormData';
const LOCAL_STORAGE_KEY_SELECTED_ID = 'stationaryAppSelectedId';
const LOCAL_STORAGE_KEY_STOCK = 'stationaryAppStock';

// Configure pdf.js worker
pdfjsLib.GlobalWorkerOptions.workerSrc = `https://esm.sh/pdfjs-dist@4.4.168/build/pdf.worker.mjs`;

// Helper functions for display logic
const formatItemsForDisplay = (items: Record<string, number> | string[]): string => {
    if (!items) return '';
    // Legacy support for old data format
    if (Array.isArray(items)) {
        return items.length > 0 ? items.join(', ') : 'N/A';
    }
    if (typeof items === 'object') {
        const entries = Object.entries(items).filter(([, quantity]) => quantity > 0);
        if (entries.length === 0) return 'N/A';
        return entries.map(([item, quantity]) => `${item} (${quantity})`).join(', ');
    }
    return 'N/A';
};

const calculateTotalItems = (items: Record<string, number> | string[]): number => {
    if (!items) return 0;
    // Legacy support for old data format
    if (Array.isArray(items)) {
        return items.length;
    }
    if (typeof items === 'object') {
        return Object.values(items).reduce((sum, quantity) => sum + (Number(quantity) || 0), 0);
    }
    return 0;
};


// Helper components defined outside the main component to prevent re-creation on re-renders.

interface CustomButtonProps {
    onClick: () => void;
    disabled?: boolean;
    color: 'blue' | 'green' | 'red' | 'gray' | 'black';
    children: React.ReactNode;
    isIconOnly?: boolean;
    title?: string;
}

const CustomButton: React.FC<CustomButtonProps> = ({ onClick, disabled = false, color, children, isIconOnly = false, title }) => {
    const colorClasses = {
        blue: 'bg-indigo-600 hover:bg-indigo-700 focus:ring-indigo-500',
        green: 'bg-green-600 hover:bg-green-700 focus:ring-green-500',
        red: 'bg-red-600 hover:bg-red-700 focus:ring-red-500',
        gray: 'bg-gray-500 hover:bg-gray-600 focus:ring-gray-500',
        black: 'bg-black hover:bg-gray-800 focus:ring-gray-700',
    };
    const paddingClasses = isIconOnly ? 'p-3' : 'px-6 py-2';
    return (
        <button
            type="button"
            onClick={onClick}
            disabled={disabled}
            title={title}
            className={`flex items-center justify-center ${paddingClasses} text-white font-semibold rounded-full shadow-md transition-colors duration-200 ease-in-out focus:outline-none focus:ring-2 focus:ring-offset-2 ${colorClasses[color]} ${disabled ? 'opacity-50 cursor-not-allowed' : ''}`}
        >
            {children}
        </button>
    );
};


interface ConfirmationModalProps {
    isOpen: boolean;
    onConfirm: () => void;
    onCancel: () => void;
    title: string;
    children: React.ReactNode;
    confirmButtonText?: string;
}

const ConfirmationModal: React.FC<ConfirmationModalProps> = ({ isOpen, onConfirm, onCancel, title, children, confirmButtonText = 'Confirm' }) => {
    useEffect(() => {
        const handleEsc = (event: KeyboardEvent) => {
            if (event.key === 'Escape') {
                onCancel();
            }
        };
        if (isOpen) {
            window.addEventListener('keydown', handleEsc);
        }

        return () => {
            window.removeEventListener('keydown', handleEsc);
        };
    }, [isOpen, onCancel]);

    if (!isOpen) return null;

    return (
        <div 
            className="fixed inset-0 bg-black/60 z-50 flex justify-center items-center p-4" 
            aria-labelledby="modal-title" 
            role="dialog" 
            aria-modal="true"
            onClick={onCancel}
        >
            <div 
                className="bg-white rounded-xl shadow-2xl p-6 sm:p-8 w-full max-w-md transform transition-all duration-300 scale-95 opacity-0 animate-fade-in-scale"
                onClick={e => e.stopPropagation()} // Prevent closing when clicking inside the modal
            >
                <div className="sm:flex sm:items-start">
                    <div className="mx-auto flex-shrink-0 flex items-center justify-center h-12 w-12 rounded-full bg-red-100 sm:mx-0 sm:h-10 sm:w-10">
                        <svg className="h-6 w-6 text-red-600" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth="2" stroke="currentColor" aria-hidden="true">
                            <path strokeLinecap="round" strokeLinejoin="round" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z" />
                        </svg>
                    </div>
                    <div className="mt-3 text-center sm:mt-0 sm:ml-4 sm:text-left">
                        <h3 className="text-xl leading-6 font-bold text-gray-900" id="modal-title">{title}</h3>
                        <div className="mt-2">
                            <div className="text-sm text-gray-600">{children}</div>
                        </div>
                    </div>
                </div>
                <div className="mt-6 sm:mt-8 sm:flex sm:flex-row-reverse gap-3">
                    <button
                        type="button"
                        className="w-full inline-flex justify-center rounded-full border border-transparent shadow-sm px-6 py-2 bg-red-600 text-base font-medium text-white hover:bg-red-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-red-500 sm:w-auto sm:text-sm"
                        onClick={onConfirm}
                    >
                        {confirmButtonText}
                    </button>
                    <button
                        type="button"
                        className="mt-3 w-full inline-flex justify-center rounded-full border border-gray-300 shadow-sm px-6 py-2 bg-white text-base font-medium text-gray-700 hover:bg-gray-50 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-indigo-500 sm:mt-0 sm:w-auto sm:text-sm"
                        onClick={onCancel}
                    >
                        Cancel
                    </button>
                </div>
            </div>
            <style>{`
                @keyframes fade-in-scale {
                    from { transform: scale(0.95); opacity: 0; }
                    to { transform: scale(1); opacity: 1; }
                }
                .animate-fade-in-scale { animation: fade-in-scale 0.2s ease-out forwards; }
            `}</style>
        </div>
    );
};


export default function App() {
    const [reports, setReports] = useState<Report[]>(() => {
        try {
            const savedReportsJSON = window.localStorage.getItem(LOCAL_STORAGE_KEY_REPORTS);
            if (savedReportsJSON) {
                const savedReports = JSON.parse(savedReportsJSON);
                if (Array.isArray(savedReports)) {
                    // Data migration for backward compatibility
                    return savedReports.map((report: any) => {
                        const newReport = { ...report };
                        // Migrate items from string[] to Record<string, number>
                        if (Array.isArray(newReport.items)) {
                            const newItems: Record<string, number> = {};
                            newReport.items.forEach((item: string) => {
                                newItems[item] = (newItems[item] || 0) + 1;
                            });
                            newReport.items = newItems;
                        } else if (!newReport.items || typeof newReport.items !== 'object') {
                            newReport.items = {};
                        }
                        // Add default status if missing
                        if (newReport.status !== 'Done') {
                            newReport.status = 'Process';
                        }
                        return newReport;
                    });
                }
            }
        } catch (error) {
            console.error("Error reading reports from localStorage:", error);
        }
        return initialReports;
    });

    const [formData, setFormData] = useState<Omit<Report, 'id'>>(() => {
        try {
            const savedFormDataJSON = window.localStorage.getItem(LOCAL_STORAGE_KEY_FORM_DATA);
            if (savedFormDataJSON) {
                const savedFormData = JSON.parse(savedFormDataJSON);
                if (typeof savedFormData === 'object' && savedFormData !== null && 'requesterName' in savedFormData) {
                    // Migrate legacy form data
                    if (Array.isArray(savedFormData.items)) {
                        const newItems: Record<string, number> = {};
                        savedFormData.items.forEach((item: string) => {
                            newItems[item] = (newItems[item] || 0) + 1;
                        });
                        savedFormData.items = newItems;
                    }
                    if (!savedFormData.items || typeof savedFormData.items !== 'object') {
                       savedFormData.items = {};
                    }
                    if (savedFormData.status !== 'Done') {
                        savedFormData.status = 'Process';
                    }
                    return savedFormData;
                }
            }
        } catch (error) {
            console.error("Error reading form data from localStorage:", error);
        }
        return initialFormData;
    });

    const [selectedReportId, setSelectedReportId] = useState<string | null>(() => {
        try {
            const savedIdJSON = window.localStorage.getItem(LOCAL_STORAGE_KEY_SELECTED_ID);
            if (savedIdJSON) {
                return JSON.parse(savedIdJSON);
            }
        } catch (error) {
            console.error("Error reading selected report ID from localStorage:", error);
        }
        return null;
    });

    const [stock, setStock] = useState<Record<string, StockItem>>(() => {
        const allItems = new Set([...STATIONARY_ITEMS_ROW1, ...STATIONARY_ITEMS_ROW2]);
        let savedStock: Record<string, any> = {};
    
        try {
            const savedStockJSON = window.localStorage.getItem(LOCAL_STORAGE_KEY_STOCK);
            if (savedStockJSON) {
                savedStock = JSON.parse(savedStockJSON);
            }
        } catch (error) {
            console.error("Error reading stock from localStorage:", error);
        }
        
        const migratedStock: Record<string, StockItem> = {};
        const today = new Date().toISOString().split('T')[0];
    
        allItems.forEach(item => {
            const stockItem = savedStock[item];
            if (stockItem !== undefined) {
                if (typeof stockItem === 'number') {
                    // Migrate from: number
                    migratedStock[item] = { quantity: stockItem, lastInDate: today, lastOutDate: '', lastUpdateQuantity: 0 };
                } else if (typeof stockItem === 'object' && stockItem !== null && 'dateAdded' in stockItem) {
                    // Migrate from: { quantity, dateAdded }
                    migratedStock[item] = { 
                        quantity: Number(stockItem.quantity) || 0,
                        lastInDate: stockItem.dateAdded || today,
                        lastOutDate: '', // Add new field
                        lastUpdateQuantity: 0
                    };
                } else if (typeof stockItem === 'object' && stockItem !== null && 'quantity' in stockItem && 'lastInDate' in stockItem) {
                    // Already in the current format or newer, just ensure all fields exist
                     migratedStock[item] = {
                        quantity: Number(stockItem.quantity) || 0,
                        lastInDate: stockItem.lastInDate || '',
                        lastOutDate: stockItem.lastOutDate || '',
                        lastUpdateQuantity: Number(stockItem.lastUpdateQuantity) || 0
                     };
                } else {
                    // Invalid data, initialize fresh
                    migratedStock[item] = { quantity: 0, lastInDate: '', lastOutDate: '', lastUpdateQuantity: 0 };
                }
            } else {
                // New item not in storage, initialize
                migratedStock[item] = { quantity: 0, lastInDate: '', lastOutDate: '', lastUpdateQuantity: 0 };
            }
        });
        
        return migratedStock;
    });

    const [campusFilter, setCampusFilter] = useState('');
    const [descriptionFilter, setDescriptionFilter] = useState('');
    const [selectedMonth, setSelectedMonth] = useState<string>('');
    const [selectedWeek, setSelectedWeek] = useState<string>('');
    const [saveStatus, setSaveStatus] = useState<'idle' | 'saved'>('idle');
    const [isConfirmingDelete, setIsConfirmingDelete] = useState(false);
    const saveTimeoutRef = useRef<number | null>(null);
    const reportsDidMount = useRef(false);
    const stockDidMount = useRef(false);

    // Stock Management State
    const [isEditingStock, setIsEditingStock] = useState(false);
    const [tempStock, setTempStock] = useState<Record<string, StockItem>>(stock);
    const [isConfirmingClearStock, setIsConfirmingClearStock] = useState(false);


    // PDF Import State
    const [isImporting, setIsImporting] = useState(false);
    const importFileRef = useRef<HTMLInputElement>(null);


    const triggerSaveStatus = useCallback(() => {
        if (saveTimeoutRef.current) {
            clearTimeout(saveTimeoutRef.current);
        }
        setSaveStatus('saved');
        saveTimeoutRef.current = window.setTimeout(() => {
            setSaveStatus('idle');
        }, 2000);
    }, []);

    useEffect(() => {
        if (reportsDidMount.current) {
            try {
                window.localStorage.setItem(LOCAL_STORAGE_KEY_REPORTS, JSON.stringify(reports));
                triggerSaveStatus();
            } catch (error) {
                console.error("Error saving reports to localStorage:", error);
                alert("Error: Could not save your reports. Your browser's local storage may be full or disabled. Please check your browser settings and try again.");
            }
        } else {
            reportsDidMount.current = true;
        }
    }, [reports, triggerSaveStatus]);

    useEffect(() => {
        if (stockDidMount.current) {
            try {
                window.localStorage.setItem(LOCAL_STORAGE_KEY_STOCK, JSON.stringify(stock));
            } catch (error) {
                console.error("Error saving stock to localStorage:", error);
            }
        } else {
            stockDidMount.current = true;
        }
    }, [stock]);

    useEffect(() => {
        if (!isEditingStock) {
            setTempStock(stock);
        }
    }, [stock, isEditingStock]);

    useEffect(() => {
        try {
            window.localStorage.setItem(LOCAL_STORAGE_KEY_FORM_DATA, JSON.stringify(formData));
        } catch (error) {
            console.error("Error saving form data to localStorage:", error);
        }
    }, [formData]);

    useEffect(() => {
        try {
            window.localStorage.setItem(LOCAL_STORAGE_KEY_SELECTED_ID, JSON.stringify(selectedReportId));
        } catch (error) {
            console.error("Error saving selected report ID to localStorage:", error);
        }
    }, [selectedReportId]);

    const handleInputChange = useCallback((e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement>) => {
        const { name, value } = e.target;
        setFormData(prev => ({ ...prev, [name]: value as any }));
    }, []);

    const handleItemQuantityChange = useCallback((item: string, value: string) => {
        const quantity = parseInt(value, 10);
        setFormData(prev => {
            const newItems = { ...prev.items };
            if (!isNaN(quantity) && quantity > 0) {
                newItems[item] = quantity;
            } else {
                delete newItems[item]; // Remove item if quantity is 0, empty or invalid
            }
            return { ...prev, items: newItems };
        });
    }, []);

    const clearForm = useCallback(() => {
        setFormData(initialFormData);
        setSelectedReportId(null);
    }, []);

    const handleAddReport = useCallback(() => {
        if (!formData.requesterName || !formData.campus || !formData.importDate || !formData.exportDate) {
            alert("Please fill all fields, including dates.");
            return;
        }
    
        // --- Stock Logic ---
        if (formData.status === 'Done') {
            const itemsToDeduct = formData.items;
            const insufficientItems: string[] = [];
            for (const [item, quantity] of Object.entries(itemsToDeduct)) {
                if ((stock[item]?.quantity || 0) < quantity) {
                    insufficientItems.push(`${item} (requested ${quantity}, available ${stock[item]?.quantity || 0})`);
                }
            }
    
            if (insufficientItems.length > 0) {
                alert(`Cannot add report. Insufficient stock for: ${insufficientItems.join(', ')}.`);
                return; // Block the action
            }
    
            // Deduct from stock
            setStock(prevStock => {
                const newStock = JSON.parse(JSON.stringify(prevStock)); // Deep copy
                const today = new Date().toISOString().split('T')[0];
                for (const [item, quantity] of Object.entries(itemsToDeduct)) {
                    const numQuantity = Number(quantity) || 0;
                    if (newStock[item]) {
                        newStock[item].quantity = (newStock[item].quantity || 0) - numQuantity;
                        newStock[item].lastOutDate = today;
                        newStock[item].lastUpdateQuantity = -numQuantity;
                    }
                }
                return newStock;
            });
        }
        // --- End Stock Logic ---
    
        const newReport: Report = {
            id: new Date().toISOString(),
            ...formData,
        };
        setReports(prev => [...prev, newReport]);
        clearForm();
    }, [formData, clearForm, stock]);

    const handleSelectReport = useCallback((report: Report) => {
        setSelectedReportId(report.id);
        setFormData({
            requesterName: report.requesterName,
            campus: report.campus,
            importDate: report.importDate,
            exportDate: report.exportDate,
            items: report.items && typeof report.items === 'object' && !Array.isArray(report.items) 
                   ? { ...report.items }
                   : {}, // The migration should prevent array items, but this is a safeguard.
            status: report.status || 'Process'
        });
    }, []);

    const handleUpdateReport = useCallback(() => {
        if (!selectedReportId) return;
    
        const originalReport = reports.find(r => r.id === selectedReportId);
        if (!originalReport) return;
    
        const updatedData = formData;
        const insufficientItems: string[] = [];
        const originalItems = originalReport.items || {};
        const updatedItems = updatedData.items || {};
    
        // --- Stock Check Logic (Run this before any state updates) ---
        if (originalReport.status === 'Process' && updatedData.status === 'Done') {
            // Case 1: Transitioning from Process to Done. Check all new items against current stock.
            for (const [item, quantity] of Object.entries(updatedItems)) {
                if ((stock[item]?.quantity || 0) < quantity) {
                    insufficientItems.push(`${item} (requested ${quantity}, available ${stock[item]?.quantity || 0})`);
                }
            }
        } else if (originalReport.status === 'Done' && updatedData.status === 'Done') {
            // Case 2: Staying as Done. Check only the *additional* items being requested (the delta).
            const allItems = new Set([...Object.keys(originalItems), ...Object.keys(updatedItems)]);
            allItems.forEach(item => {
                const originalQty = Number(originalItems[item] || 0);
                const updatedQty = Number(updatedItems[item] || 0);
                const delta = updatedQty - originalQty; // A positive delta means more items are being taken.
    
                if (delta > 0 && (stock[item]?.quantity || 0) < delta) {
                     insufficientItems.push(`${item} (requested ${delta} more, available ${stock[item]?.quantity || 0})`);
                }
            });
        }
    
        if (insufficientItems.length > 0) {
            alert(`Cannot update report. Insufficient stock for: ${insufficientItems.join(', ')}.`);
            return; // Block the update
        }
        // --- End Stock Check Logic ---
    
        // --- Stock Update Logic ---
        setStock(prevStock => {
            const newStock = JSON.parse(JSON.stringify(prevStock)); // Deep copy
            const today = new Date().toISOString().split('T')[0];
    
            // Case 1: Process -> Done (Deduct new items)
            if (originalReport.status === 'Process' && updatedData.status === 'Done') {
                for (const [item, quantity] of Object.entries(updatedItems)) {
                    const numQuantity = Number(quantity) || 0;
                    if (newStock[item]) {
                         newStock[item].quantity -= numQuantity;
                         newStock[item].lastOutDate = today;
                         newStock[item].lastUpdateQuantity = -numQuantity;
                    }
                }
            }
            // Case 2: Done -> Process (Add back old items)
            else if (originalReport.status === 'Done' && updatedData.status === 'Process') {
                for (const [item, quantity] of Object.entries(originalItems)) {
                    const numQuantity = Number(quantity) || 0;
                    if (newStock[item]) {
                        newStock[item].quantity += numQuantity;
                        newStock[item].lastInDate = today;
                        newStock[item].lastUpdateQuantity = numQuantity;
                    }
                }
            }
            // Case 3: Done -> Done (Calculate delta)
            else if (originalReport.status === 'Done' && updatedData.status === 'Done') {
                const allItems = new Set([...Object.keys(originalItems), ...Object.keys(updatedItems)]);
                allItems.forEach(item => {
                    const originalQty = Number(originalItems[item] || 0);
                    const updatedQty = Number(updatedItems[item] || 0);
                    const delta = originalQty - updatedQty; // positive: items returned, negative: more items taken
                    if (newStock[item] && delta !== 0) {
                        newStock[item].quantity += delta;
                        newStock[item].lastUpdateQuantity = delta;
                        if (delta > 0) { // Items returned to stock
                            newStock[item].lastInDate = today;
                        } else { // More items taken from stock
                            newStock[item].lastOutDate = today;
                        }
                    }
                });
            }
            // Case 4: Process -> Process (No stock change)
            
            return newStock;
        });
        // --- End Stock Update Logic ---
    
        setReports(prev =>
            prev.map(r =>
                r.id === selectedReportId ? { ...r, ...updatedData } : r
            )
        );
        clearForm();
    }, [selectedReportId, formData, clearForm, reports, stock]);
    
    const handleConfirmDelete = useCallback(() => {
        if (!selectedReportId) return;
    
        // --- Stock Logic ---
        const reportToDelete = reports.find(r => r.id === selectedReportId);
        if (reportToDelete && reportToDelete.status === 'Done') {
            const itemsToAddBack = reportToDelete.items;
            setStock(prevStock => {
                const newStock = JSON.parse(JSON.stringify(prevStock)); // Deep copy
                const today = new Date().toISOString().split('T')[0];
                for (const [item, quantity] of Object.entries(itemsToAddBack)) {
                    const numQuantity = Number(quantity) || 0;
                    if (newStock[item]) {
                        newStock[item].quantity += numQuantity;
                        newStock[item].lastInDate = today;
                        newStock[item].lastUpdateQuantity = numQuantity;
                    }
                }
                return newStock;
            });
        }
        // --- End Stock Logic ---
    
        setReports(prev => prev.filter(r => r.id !== selectedReportId));
        clearForm();
        setIsConfirmingDelete(false);
    }, [selectedReportId, clearForm, reports]);
    
    const handleDeleteReport = useCallback(() => {
        if (!selectedReportId) return;
        setIsConfirmingDelete(true);
    }, [selectedReportId]);

    const isEditing = selectedReportId !== null;

    const availableMonths = useMemo(() => {
        const months = new Set<string>();
        reports.forEach(report => {
            if (report.importDate) {
                const month = report.importDate.substring(0, 7); // 'YYYY-MM'
                months.add(month);
            }
        });
        return Array.from(months).sort().reverse();
    }, [reports]);

    const getStartOfWeek = (dateString: string): Date => {
        const date = new Date(dateString + 'T00:00:00');
        const day = date.getDay();
        const diff = date.getDate() - day + (day === 0 ? -6 : 1); // adjust when day is sunday
        return new Date(date.setDate(diff));
    };

    const availableWeeks = useMemo(() => {
        const weeks = new Map<string, string>(); // Map of 'YYYY-MM-DD' -> 'Week of Mon DD, YYYY'
        const reportsToProcess = selectedMonth
            ? reports.filter(r => r.importDate.startsWith(selectedMonth))
            : reports;

        reportsToProcess.forEach(report => {
            if (report.importDate) {
                try {
                    const startOfWeek = getStartOfWeek(report.importDate);
                    const startOfWeekISO = startOfWeek.toISOString().split('T')[0];
                    if (!weeks.has(startOfWeekISO)) {
                        const weekLabel = `Week of ${startOfWeek.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' })}`;
                        weeks.set(startOfWeekISO, weekLabel);
                    }
                } catch(e) {
                    // Ignore invalid dates
                }
            }
        });

        return Array.from(weeks.entries())
            .map(([value, label]) => ({ value, label }))
            .sort((a, b) => b.value.localeCompare(a.value)); // Sort descending
    }, [reports, selectedMonth]);

    const handleMonthChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
        setSelectedMonth(e.target.value);
        setSelectedWeek(''); // Reset week when month changes
    };

    const handleWeekChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
        setSelectedWeek(e.target.value);
    };
    
    const filteredReports = useMemo(() => reports.filter(report => {
        const campusMatch = campusFilter ? report.campus === campusFilter : true;
        const descriptionMatch = descriptionFilter
            ? formatItemsForDisplay(report.items).toLowerCase().includes(descriptionFilter.toLowerCase())
            : true;

        let dateMatch = true;
        if (selectedWeek) {
            try {
                const startOfWeek = new Date(selectedWeek + 'T00:00:00');
                const endOfWeek = new Date(startOfWeek);
                endOfWeek.setDate(startOfWeek.getDate() + 7);
                const reportDate = new Date(report.importDate + 'T00:00:00');
                dateMatch = reportDate >= startOfWeek && reportDate < endOfWeek;
            } catch (e) {
                dateMatch = false;
            }
        } else if (selectedMonth) {
            dateMatch = report.importDate.startsWith(selectedMonth);
        }
        
        return campusMatch && descriptionMatch && dateMatch;
    }), [reports, campusFilter, selectedMonth, selectedWeek, descriptionFilter]);

    const itemCounts = useMemo(() => {
        return filteredReports.reduce((acc, report) => {
            if (report.items && typeof report.items === 'object') {
                for (const [item, quantity] of Object.entries(report.items)) {
                    acc[item] = (acc[item] || 0) + (Number(quantity) || 0);
                }
            }
            return acc;
        }, {} as Record<string, number>);
    }, [filteredReports]);
    
    const reportToDelete = useMemo(() =>
        selectedReportId ? reports.find(r => r.id === selectedReportId) : null,
        [reports, selectedReportId]
    );

    const handleExportPDF = useCallback(() => {
        if (filteredReports.length === 0 && Object.keys(stock).length === 0) {
            alert("No data to export.");
            return;
        }
    
        const doneReports = filteredReports.filter(r => (r.status || 'Process') === 'Done');
        const processReports = filteredReports.filter(r => (r.status || 'Process') === 'Process');
    
        const calculateItemCounts = (reports: Report[]) => {
            return reports.reduce((acc, report) => {
                if (report.items && typeof report.items === 'object') {
                    for (const [item, quantity] of Object.entries(report.items)) {
                        acc[item] = (acc[item] || 0) + (Number(quantity) || 0);
                    }
                }
                return acc;
            }, {} as Record<string, number>);
        };
    
        const itemCountsDone = calculateItemCounts(doneReports);
        const itemCountsProcess = calculateItemCounts(processReports);
        const itemCountsTotal = calculateItemCounts(filteredReports); // Grand total
    
        const doc = new jsPDF();
        
        let periodName = 'All Time';
        if (selectedWeek) {
            const weekData = availableWeeks.find(w => w.value === selectedWeek);
            periodName = weekData ? weekData.label : `Week of ${selectedWeek}`;
        } else if (selectedMonth) {
            periodName = new Date(selectedMonth + '-02').toLocaleString('en-US', { month: 'long', year: 'numeric' });
        }
        const campusName = campusFilter || 'All Campuses';
    
        doc.setFontSize(18);
        doc.text('Stationary Report', 14, 22);
        doc.setFontSize(12);
        doc.text(`Campus: ${campusName}`, 14, 30);
        doc.text(`Period: ${periodName}`, 14, 36);
    
        let currentY = 45;

        // --- Overall Summary Section ---
        if (Object.keys(itemCountsTotal).length > 0) {
            doc.setFontSize(16);
            doc.setTextColor(45, 55, 72);
            doc.text('Overall Summary', 14, currentY);
            currentY += 8;
    
            doc.setFontSize(14);
            doc.setTextColor(0, 0, 0);
            doc.text('Grand Total (All Statuses)', 14, currentY);
            currentY += 7;
            doc.setFontSize(10);
            const summaryText = Object.entries(itemCountsTotal).map(([item, count]) => `${item}: ${count}`).join(' | ');
            const splitSummary = doc.splitTextToSize(summaryText, 180);
            doc.text(splitSummary, 14, currentY);
            currentY += (splitSummary.length * 4) + 5;
        }
    
        const addSectionToPdf = (title: string, reports: Report[], counts: Record<string, number>, startY: number): number => {
            if (reports.length === 0) {
                return startY;
            }
    
            if (startY > 45) { // Check ensures it's not the very first section on the page
                startY += 10;
            }
    
            doc.setFontSize(16);
            doc.setTextColor(45, 55, 72);
            doc.text(title, 14, startY);
            startY += 8;
            
            if (Object.keys(counts).length > 0) {
                doc.setFontSize(14);
                doc.setTextColor(0, 0, 0);
                doc.text('Summary (Total Items)', 14, startY);
                startY += 7;
                doc.setFontSize(10);
                const summaryText = Object.entries(counts).map(([item, count]) => `${item}: ${count}`).join(' | ');
                const splitSummary = doc.splitTextToSize(summaryText, 180);
                doc.text(splitSummary, 14, startY);
                startY += (splitSummary.length * 4) + 5;
            }
    
            const tableColumns = ["Requester Name", "Campus", "Import Date", "Export Date", "Description", "Total"];
            const tableRows = reports.map(report => [
                report.requesterName,
                report.campus,
                report.importDate,
                report.exportDate,
                formatItemsForDisplay(report.items),
                calculateTotalItems(report.items).toString(),
            ]);
    
            autoTable(doc, {
                head: [tableColumns],
                body: tableRows,
                startY: startY,
                theme: 'grid',
                headStyles: { fillColor: [45, 55, 72] },
            });
            
            return (doc as any).lastAutoTable.finalY;
        };
    
        currentY = addSectionToPdf('Status: Done', doneReports, itemCountsDone, currentY);
        currentY = addSectionToPdf('Status: Process', processReports, itemCountsProcess, currentY);
    
        // --- Stock Inventory Section ---
        let lastY = (doc as any).lastAutoTable.finalY || currentY;
        if (lastY > 250) { // Check if new page is needed
            doc.addPage();
            lastY = 20;
        } else {
            lastY += 15;
        }
    
        doc.setFontSize(16);
        doc.setTextColor(45, 55, 72);
        doc.text('Current Stock Inventory', 14, lastY);
        lastY += 8;

        const stockTableColumns = ["Item", "Quantity in Stock", "Date Added"];
        const stockTableRows = Object.entries(stock)
            .sort(([a], [b]) => a.localeCompare(b))
            .map(([item, { quantity, lastInDate }]) => [
                item,
                quantity.toString(),
                lastInDate || 'N/A'
            ]);
        
        autoTable(doc, {
            head: [stockTableColumns],
            body: stockTableRows,
            startY: lastY,
            theme: 'grid',
            headStyles: { fillColor: [80, 80, 80] },
        });

        const fileName = `Stationary_Report_${campusName.replace(/ /g, '_')}_${periodName.replace(/ /g, '_')}.pdf`;
        doc.save(fileName);
    }, [filteredReports, campusFilter, selectedMonth, selectedWeek, stock, availableWeeks]);
    
    const handleTriggerPdfImport = useCallback(() => {
        importFileRef.current?.click();
    }, []);

    const handlePdfImport = useCallback(async (e: React.ChangeEvent<HTMLInputElement>) => {
        const file = e.target.files?.[0];
        if (!file) return;
    
        setIsImporting(true);
    
        try {
            const arrayBuffer = await file.arrayBuffer();
    
            // Extract text from PDF
            const pdf = await pdfjsLib.getDocument(arrayBuffer).promise;
            const numPages = pdf.numPages;
            let fullText = '';
            for (let i = 1; i <= numPages; i++) {
                const page = await pdf.getPage(i);
                const textContent = await page.getTextContent();
                const pageText = textContent.items.map((item: any) => item.str).join(' ');
                fullText += pageText + '\n\n';
            }
    
            if (!fullText.trim()) {
                throw new Error("Could not extract any text from the PDF.");
            }
    
            // Use AI to parse the text
            const ai = new GoogleGenAI({ apiKey: process.env.API_KEY! });
            const prompt = `
                The following text was extracted from a PDF file. It contains one or more tables for stationary reports and a table for stock inventory.
                Please parse this text and convert it into a single JSON object.
                This JSON object must have two top-level keys: "reports" and "stock".

                1.  The "reports" key should contain a JSON array of report objects. Each object in the array should have these properties: "requesterName", "campus", "importDate", "exportDate", "items", and "status".
                    -   "requesterName": string
                    -   "campus": string
                    -   "importDate": string, in "YYYY-MM-DD" format.
                    -   "exportDate": string, in "YYYY-MM-DD" format.
                    -   "status": string, either "Process" or "Done".
                    -   "items": an object where keys are item names (string) and values are quantities (number). A description "Bk (5), Card (2)" should become \`{ "Bk": 5, "Card": 2 }\`. Empty descriptions result in an empty object {}.

                2.  The "stock" key should contain a JSON object representing the stock inventory from the "Current Stock Inventory" table. This table only has "Item", "Quantity in Stock", and "Date Added" columns.
                    -   The keys of this object should be the item names (string).
                    -   The values should be objects with two properties: "quantity" (number) and "lastInDate" (string, in "YYYY-MM-DD" format or "N/A").
                    -   The "lastOutDate" is not present in the source table, so do not include it in the output.
                    -   Example: \`{ "Bk": { "quantity": 100, "lastInDate": "2024-05-10" } }\`

                Here is the text to parse:
                ---
                ${fullText}
                ---

                Return ONLY the single JSON object, without any surrounding text or markdown.
            `;

            const response = await ai.models.generateContent({
                model: 'gemini-2.5-flash',
                contents: prompt,
                config: {
                    responseMimeType: "application/json",
                },
            });

            let jsonStr = response.text.trim();
            const fenceRegex = /^```(\w*)?\s*\n?(.*?)\n?\s*```$/s;
            const match = jsonStr.match(fenceRegex);
            if (match && match[2]) {
              jsonStr = match[2].trim();
            }
            
            const parsedData = JSON.parse(jsonStr);

            if (!parsedData || typeof parsedData !== 'object') {
                throw new Error("AI response is not a valid JSON object.");
            }
            
            // Process Reports
            const parsedReports = parsedData.reports;
            if (!Array.isArray(parsedReports)) {
                throw new Error("AI did not return a valid array of reports in the 'reports' key.");
            }
    
            const newReports: Report[] = parsedReports.map((item: any) => {
                const status: 'Process' | 'Done' = item.status === 'Done' ? 'Done' : 'Process';
                const items: Record<string, number> = item.items && typeof item.items === 'object' && !Array.isArray(item.items) ? item.items : {};
                
                return {
                    ...initialFormData,
                    id: `imported-${new Date().toISOString()}-${Math.random()}`,
                    requesterName: item.requesterName || '',
                    campus: item.campus || '',
                    importDate: item.importDate || '',
                    exportDate: item.exportDate || '',
                    items: items,
                    status: status,
                }
            }).filter(r => r.requesterName && r.campus && r.importDate);
    
            // Process Stock
            const parsedStock = parsedData.stock;
            if (!parsedStock || typeof parsedStock !== 'object') {
                throw new Error("AI did not return a valid object in the 'stock' key.");
            }

            const newStock: Record<string, StockItem> = {};

            for (const item in parsedStock) {
                if (Object.prototype.hasOwnProperty.call(parsedStock, item)) {
                    const stockItem = parsedStock[item];
                    if (stockItem && typeof stockItem.quantity === 'number') {
                       newStock[item] = {
                           quantity: stockItem.quantity,
                           lastInDate: (typeof stockItem.lastInDate === 'string' && stockItem.lastInDate !== 'N/A') ? stockItem.lastInDate : '',
                           lastOutDate: '', // Reset on import as it's not in the PDF
                           lastUpdateQuantity: 0,
                       };
                    }
                }
            }

            setReports(newReports);
            setStock(newStock);
            alert(`Successfully imported ${newReports.length} reports and replaced the stock inventory.`);
    
        } catch (error) {
            console.error("Failed to import PDF:", error);
            const errorMessage = error instanceof Error ? error.message : "An unknown error occurred.";
            alert(`Error importing PDF: ${errorMessage}`);
        } finally {
            setIsImporting(false);
            if (e.target) {
                e.target.value = ''; // Reset file input
            }
        }
    }, [setReports, setStock]);

    const handleTempStockChange = useCallback((item: string, value: string) => {
        const quantity = parseInt(value, 10);
        setTempStock(prev => {
            const currentItem = prev[item] || { quantity: 0, lastInDate: '', lastOutDate: '', lastUpdateQuantity: 0 };
            return {
                ...prev,
                [item]: {
                    ...currentItem,
                    quantity: isNaN(quantity) || quantity < 0 ? 0 : quantity
                }
            };
        });
    }, []);

    const handleSaveStock = useCallback(() => {
        const newStock = JSON.parse(JSON.stringify(tempStock));
        const today = new Date().toISOString().split('T')[0];
    
        // Determine date changes and last update quantity by comparing with original stock
        for (const item in newStock) {
            const oldStockItem = stock[item] || { quantity: 0, lastInDate: '', lastOutDate: '', lastUpdateQuantity: 0 };
            const oldQuantity = oldStockItem.quantity;
            const newQuantity = newStock[item].quantity;
            const delta = newQuantity - oldQuantity;
    
            if (delta !== 0) {
                newStock[item].lastUpdateQuantity = delta;
                if (delta > 0) {
                    newStock[item].lastInDate = today;
                    newStock[item].lastOutDate = oldStockItem.lastOutDate; // Preserve old date
                } else {
                    newStock[item].lastOutDate = today;
                    newStock[item].lastInDate = oldStockItem.lastInDate; // Preserve old date
                }
            } else {
                // Quantities are the same, preserve original data
                newStock[item].lastInDate = oldStockItem.lastInDate;
                newStock[item].lastOutDate = oldStockItem.lastOutDate;
                newStock[item].lastUpdateQuantity = oldStockItem.lastUpdateQuantity;
            }
        }
    
        setStock(newStock);
        setIsEditingStock(false);
        triggerSaveStatus();
    }, [tempStock, stock, triggerSaveStatus]);

    const handleCancelEditStock = useCallback(() => {
        setTempStock(stock); // Revert changes
        setIsEditingStock(false);
    }, [stock]);

    const handleConfirmClearStock = useCallback(() => {
        setStock(prevStock => {
            const clearedStock: Record<string, StockItem> = {};
            Object.keys(prevStock).forEach(itemKey => {
                clearedStock[itemKey] = { quantity: 0, lastInDate: '', lastOutDate: '', lastUpdateQuantity: 0 };
            });
            return clearedStock;
        });
        setIsConfirmingClearStock(false);
        triggerSaveStatus();
    }, [triggerSaveStatus]);

    const handleExportStockPDF = useCallback(() => {
        if (Object.keys(stock).length === 0) {
            alert("No stock data to export.");
            return;
        }

        const doc = new jsPDF();
        const today = new Date();
        const formattedDate = today.toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });
        const formattedTime = today.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' });

        doc.setFontSize(18);
        doc.text('Stock Inventory Report', 14, 22);
        doc.setFontSize(12);
        doc.text(`Generated on: ${formattedDate} at ${formattedTime}`, 14, 30);

        const stockTableColumns = ["Item", "Quantity in Stock", "Last Date In", "Last Date Out"];
        const stockTableRows = Object.entries(stock)
            .sort(([a], [b]) => a.localeCompare(b))
            .map(([item, { quantity, lastInDate, lastOutDate }]) => [
                item,
                quantity.toString(),
                lastInDate || 'N/A',
                lastOutDate || 'N/A'
            ]);
        
        autoTable(doc, {
            head: [stockTableColumns],
            body: stockTableRows,
            startY: 40,
            theme: 'grid',
            headStyles: { fillColor: [45, 55, 72] }, // Dark grey header
        });

        const fileName = `Stock_Inventory_Report_${today.toISOString().split('T')[0]}.pdf`;
        doc.save(fileName);
    }, [stock]);

    return (
        <>
            <div className="min-h-screen bg-gray-100 flex items-center justify-center p-4 font-sans">
                <div className="w-full max-w-6xl mx-auto p-4 rounded-2xl">
                    <div className="bg-white rounded-lg p-6 sm:p-8">
                        <h1 className="text-4xl sm:text-5xl font-koulen text-center text-gray-800 mb-8">Report Stationary</h1>
                        
                        <form className="space-y-6">
                            {/* Requester Name and Campus */}
                            <div className="grid grid-cols-1 sm:grid-cols-3 gap-6">
                                <div className="relative">
                                    <label className="absolute -top-3 left-3 bg-white px-1 text-sm font-medium text-gray-600 font-serif-khmer">ឈ្មោះអ្នកស្នើសុំ</label>
                                    <input
                                        type="text"
                                        name="requesterName"
                                        value={formData.requesterName}
                                        onChange={handleInputChange}
                                        className="w-full px-4 py-3 border border-gray-300 rounded-xl focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 transition-colors"
                                    />
                                </div>
                                <div className="relative">
                                    <label className="absolute -top-3 left-3 bg-white px-1 text-sm font-medium text-gray-600 font-serif-khmer">សាខា</label>
                                    <select
                                        name="campus"
                                        value={formData.campus}
                                        onChange={handleInputChange}
                                        className="w-full px-4 py-3 border border-gray-300 rounded-xl focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 transition-colors bg-white appearance-none"
                                    >
                                        <option value="" disabled>Select a campus</option>
                                        {CAMPUS_OPTIONS.map(campus => (
                                            <option key={campus} value={campus}>{campus}</option>
                                        ))}
                                    </select>
                                </div>
                                <div className="relative">
                                    <label className="absolute -top-3 left-3 bg-white px-1 text-sm font-medium text-gray-600 font-serif-khmer">ស្ថានភាព</label>
                                    <select
                                        name="status"
                                        value={formData.status}
                                        onChange={handleInputChange}
                                        className="w-full px-4 py-3 border border-gray-300 rounded-xl focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 transition-colors bg-white appearance-none"
                                    >
                                        <option value="Process">Process</option>
                                        <option value="Done">Done</option>
                                    </select>
                                </div>
                            </div>

                            {/* Date Inputs */}
                            <div className="grid grid-cols-1 sm:grid-cols-2 gap-6">
                                <div className="relative">
                                    <label className="absolute -top-3 left-3 bg-white px-1 text-sm font-medium text-gray-600 font-serif-khmer">ថ្ងៃនាំចូល</label>
                                    <input
                                        type="date"
                                        name="importDate"
                                        value={formData.importDate}
                                        onChange={handleInputChange}
                                        className="w-full px-4 py-3 border border-gray-300 rounded-xl focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 transition-colors"
                                    />
                                </div>
                                <div className="relative">
                                    <label className="absolute -top-3 left-3 bg-white px-1 text-sm font-medium text-gray-600 font-serif-khmer">ថ្ងៃនាំចេញ</label>
                                    <input
                                        type="date"
                                        name="exportDate"
                                        value={formData.exportDate}
                                        onChange={handleInputChange}
                                        className="w-full px-4 py-3 border border-gray-300 rounded-xl focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 transition-colors"
                                    />
                                </div>
                            </div>

                            {/* Item Quantities */}
                            <div className="space-y-4 pt-2">
                                <label className="text-base font-medium text-gray-800 font-serif-khmer">សម្ភារៈ</label>
                                <div className="grid grid-cols-2 sm:grid-cols-4 gap-x-8 gap-y-4 p-4 border border-gray-200 rounded-lg">
                                    {[...STATIONARY_ITEMS_ROW1, ...STATIONARY_ITEMS_ROW2].map(item => (
                                        <div key={item} className="flex items-center justify-between">
                                            <label htmlFor={`item-${item}`} className="text-gray-700 font-medium">{item}</label>
                                            <input
                                                id={`item-${item}`}
                                                type="number"
                                                min="0"
                                                placeholder="0"
                                                value={formData.items[item] || ''}
                                                onChange={(e) => handleItemQuantityChange(item, e.target.value)}
                                                className="w-20 px-3 py-1.5 border border-gray-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 transition-colors text-center"
                                                aria-label={`Quantity for ${item}`}
                                            />
                                        </div>
                                    ))}
                                </div>
                            </div>
                        </form>

                        {/* Action Buttons */}
                        <div className="flex flex-wrap justify-start items-center gap-4 mt-8 mb-4">
                            <CustomButton onClick={isEditing ? handleUpdateReport : handleAddReport} color={isEditing ? 'green' : 'blue'}>
                                {isEditing ? 'Update' : 'Add'}
                            </CustomButton>
                            <CustomButton onClick={handleDeleteReport} disabled={!isEditing} color="red">
                                Delete
                            </CustomButton>
                            <CustomButton onClick={clearForm} color="gray">
                                Clear
                            </CustomButton>
                            <div className={`transition-opacity duration-300 ${saveStatus === 'saved' ? 'opacity-100' : 'opacity-0'}`}>
                                {saveStatus === 'saved' && (
                                    <div className="flex items-center text-green-600 font-semibold">
                                        <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5 mr-1" viewBox="0 0 20 20" fill="currentColor">
                                            <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clipRule="evenodd" />
                                        </svg>
                                        <span>Saved!</span>
                                    </div>
                                )}
                            </div>
                        </div>
                        
                        {/* Stock Management Section */}
                        <div className="mt-8 border-t border-gray-200 pt-8">
                            <div className="flex flex-wrap justify-between items-center gap-4 mb-4">
                                <h2 className="text-2xl font-normal text-gray-600 font-Poppins">Stock System</h2>
                                {!isEditingStock && (
                                    <div className="flex flex-wrap items-center gap-4">
                                         <CustomButton onClick={handleExportStockPDF} color="blue">
                                            <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5 mr-2 -ml-2" viewBox="0 0 20 20" fill="currentColor">
                                                <path fillRule="evenodd" d="M6 2a2 2 0 00-2 2v12a2 2 0 002 2h8a2 2 0 002-2V7.414A2 2 0 0015.414 6L12 2.586A2 2 0 0010.586 2H6zm5 6a1 1 0 10-2 0v3.586l-1.293-1.293a1 1 0 10-1.414 1.414l3 3a1 1 0 001.414 0l3-3a1 1 0 00-1.414-1.414L11 11.586V8z" clipRule="evenodd" />
                                            </svg>
                                            Export Stock
                                        </CustomButton>
                                        <CustomButton onClick={() => setIsEditingStock(true)} color="gray">
                                            Add Stock
                                        </CustomButton>
                                        <CustomButton onClick={() => setIsConfirmingClearStock(true)} color="red">
                                            Clear Stock
                                        </CustomButton>
                                    </div>
                                )}
                            </div>
                            <div className="p-4 border border-gray-200 rounded-lg">
                                {isEditingStock ? (
                                    <div className="space-y-4">
                                        <div className="grid grid-cols-2 sm:grid-cols-4 gap-x-8 gap-y-4">
                                            {[...STATIONARY_ITEMS_ROW1, ...STATIONARY_ITEMS_ROW2].map(item => (
                                                <div key={item} className="flex items-center justify-between">
                                                    <label htmlFor={`stock-item-${item}`} className="text-gray-700 font-medium">{item}</label>
                                                    <input
                                                        id={`stock-item-${item}`}
                                                        type="number"
                                                        min="0"
                                                        value={tempStock[item]?.quantity || ''}
                                                        onChange={(e) => handleTempStockChange(item, e.target.value)}
                                                        className="w-20 px-3 py-1.5 border border-gray-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 transition-colors text-center"
                                                        aria-label={`Stock quantity for ${item}`}
                                                    />
                                                </div>
                                            ))}
                                        </div>
                                        <div className="flex justify-end gap-4 mt-4">
                                            <CustomButton onClick={handleCancelEditStock} color="gray">Cancel</CustomButton>
                                            <CustomButton onClick={handleSaveStock} color="blue">Save Stock</CustomButton>
                                        </div>
                                    </div>
                                ) : (
                                    <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
                                        {Object.entries(stock).sort(([a], [b]) => a.localeCompare(b)).map(([item, { quantity, lastInDate, lastOutDate, lastUpdateQuantity }]) => (
                                            <div key={item} className="flex flex-col items-start justify-between bg-gray-50 text-gray-800 p-3 rounded-lg shadow-sm border border-gray-200 min-h-[90px]">
                                                <div className="flex items-baseline justify-between w-full">
                                                    <span className="font-koulen mr-2 text-lg">{item}</span>
                                                    <div className="flex items-baseline">
                                                        {lastUpdateQuantity !== 0 && (
                                                            <span className={`text-sm font-bold mr-2 ${lastUpdateQuantity > 0 ? 'text-green-500' : 'text-red-500'}`}>
                                                                ({lastUpdateQuantity > 0 ? '+' : ''}{lastUpdateQuantity})
                                                            </span>
                                                        )}
                                                        <span className={`font-bold text-xl ${quantity < 10 ? 'text-red-600' : 'text-green-600'}`}>{quantity}</span>
                                                    </div>
                                                </div>
                                                <div className="flex flex-col text-xs text-gray-500 mt-1 w-full text-left">
                                                    <span className="text-green-700">In: {lastInDate || 'N/A'}</span>
                                                    <span className="text-red-700">Out: {lastOutDate || 'N/A'}</span>
                                                </div>
                                            </div>
                                        ))}
                                    </div>
                                )}
                            </div>
                        </div>

                        {/* Filter & Export Controls */}
                        <div className="flex flex-col sm:flex-row sm:justify-between sm:items-end gap-4 pt-8 mt-8 border-t border-gray-200">
                            <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-6 flex-grow">
                                <div className="relative">
                                    <label className="absolute -top-3 left-3 bg-white px-1 text-sm font-medium text-gray-600 font-serif-khmer">ជ្រើសរើសសាខា</label>
                                    <select
                                        name="campusFilter"
                                        value={campusFilter}
                                        onChange={(e) => setCampusFilter(e.target.value)}
                                        className="w-full px-4 py-3 border border-gray-300 rounded-xl focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 transition-colors bg-white appearance-none"
                                    >
                                        <option value="">All Campuses</option>
                                        {CAMPUS_OPTIONS.map(campus => (
                                            <option key={campus} value={campus}>{campus}</option>
                                        ))}
                                    </select>
                                </div>
                                <div className="relative">
                                    <label className="absolute -top-3 left-3 bg-white px-1 text-sm font-medium text-gray-600 font-serif-khmer">ស្វែងរកតាមសម្ភារៈ</label>
                                    <input
                                        type="text"
                                        placeholder="Search by description..."
                                        value={descriptionFilter}
                                        onChange={(e) => setDescriptionFilter(e.target.value)}
                                        className="w-full px-4 py-3 border border-gray-300 rounded-xl focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 transition-colors"
                                    />
                                </div>
                                <div className="relative">
                                     <label className="absolute -top-3 left-3 bg-white px-1 text-sm font-medium text-gray-600 font-serif-khmer">ស្រង់តាមខែ</label>
                                     <select
                                         name="monthFilter"
                                         value={selectedMonth}
                                         onChange={handleMonthChange}
                                         className="w-full px-4 py-3 border border-gray-300 rounded-xl focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 transition-colors bg-white appearance-none"
                                     >
                                         <option value="">All Months</option>
                                         {availableMonths.map(month => (
                                             <option key={month} value={month}>
                                                 {new Date(month + '-02').toLocaleString('en-US', { month: 'long', year: 'numeric' })}
                                             </option>
                                         ))}
                                     </select>
                                </div>
                                 <div className="relative">
                                     <label className="absolute -top-3 left-3 bg-white px-1 text-sm font-medium text-gray-600 font-serif-khmer">ស្រង់តាមសប្តាហ៍</label>
                                     <select
                                         name="weekFilter"
                                         value={selectedWeek}
                                         onChange={handleWeekChange}
                                         disabled={availableWeeks.length === 0}
                                         className="w-full px-4 py-3 border border-gray-300 rounded-xl focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 transition-colors bg-white appearance-none disabled:bg-gray-100 disabled:cursor-not-allowed"
                                     >
                                         <option value="">All Weeks</option>
                                         {availableWeeks.map(({ value, label }) => (
                                             <option key={value} value={value}>{label}</option>
                                         ))}
                                     </select>
                                </div>
                            </div>
                            <div className="flex-shrink-0 pt-4 sm:pt-0 flex items-center gap-4">
                                <CustomButton onClick={handleTriggerPdfImport} color="gray" disabled={isImporting}>
                                    {isImporting ? (
                                        <>
                                            <svg className="animate-spin h-5 w-5 mr-2" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
                                                <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle>
                                                <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
                                            </svg>
                                            Importing...
                                        </>
                                    ) : (
                                        <>
                                            <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5 mr-2 inline-block" viewBox="0 0 20 20" fill="currentColor">
                                                <path d="M9.293 4.293a1 1 0 011.414 0l4 4a1 1 0 01-1.414 1.414L11 7.414V15a1 1 0 11-2 0V7.414L6.707 9.707a1 1 0 01-1.414-1.414l4-4z" />
                                                <path d="M4 11a1 1 0 011 1v3a1 1 0 001 1h8a1 1 0 001-1v-3a1 1 0 112 0v3a3 3 0 01-3 3H6a3 3 0 01-3-3v-3a1 1 0 011-1z" />
                                            </svg>
                                            Import PDF
                                        </>
                                    )}
                                </CustomButton>
                                <input
                                    type="file"
                                    ref={importFileRef}
                                    onChange={handlePdfImport}
                                    accept=".pdf"
                                    className="hidden"
                                    aria-hidden="true"
                                />
                                 <CustomButton onClick={handleExportPDF} color="blue" disabled={filteredReports.length === 0 || isImporting}>
                                    <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5 mr-2 inline-block" viewBox="0 0 20 20" fill="currentColor">
                                        <path fillRule="evenodd" d="M6 2a2 2 0 00-2 2v12a2 2 0 002 2h8a2 2 0 002-2V7.414A2 2 0 0015.414 6L12 2.586A2 2 0 0010.586 2H6zm5 6a1 1 0 10-2 0v3.586l-1.293-1.293a1 1 0 10-1.414 1.414l3 3a1 1 0 001.414 0l3-3a1 1 0 00-1.414-1.414L11 11.586V8z" clipRule="evenodd" />
                                    </svg>
                                    Export PDF
                                 </CustomButton>
                            </div>
                        </div>
                        
                        {/* Filtered Item Counts */}
                        {(campusFilter || selectedMonth || descriptionFilter) && Object.keys(itemCounts).length > 0 && (
                            <div className="mb-6 p-4 bg-gray-50 rounded-xl border border-gray-200">
                                <h3 className="text-lg font-bold text-gray-800 mb-3 font-serif-khmer">សរុបសម្ភារៈដែលបានស្មើសុំ</h3>
                                <div className="flex flex-wrap gap-3">
                                    {Object.entries(itemCounts).map(([item, count]) => (
                                        <div key={item} className="flex items-center bg-indigo-100 text-indigo-800 text-sm font-semibold px-4 py-2 rounded-full shadow-sm">
                                            <span className="font-koulen mr-2 text-base">{item}</span>
                                            <span className="font-bold text-lg">{count}</span>
                                        </div>
                                    ))}
                                </div>
                            </div>
                        )}

                        {/* Report Table */}
                        <div className="overflow-y-auto h-[300px] overflow-x-auto relative border border-gray-200 rounded-lg mt-4">
                            <table className="min-w-full bg-white">
                                <thead className="sticky top-0 bg-gray-100 z-10">
                                    <tr>
                                        <th className="py-3 px-4 text-left text-sm font-bold text-gray-600 uppercase tracking-wider font-serif-khmer">ឈ្មោះអ្នកស្នើសុំ</th>
                                        <th className="py-3 px-4 text-left text-sm font-bold text-gray-600 uppercase tracking-wider font-serif-khmer">សាខា</th>
                                        <th className="py-3 px-4 text-left text-sm font-bold text-gray-600 uppercase tracking-wider font-serif-khmer">ថ្ងៃនាំចូល</th>
                                        <th className="py-3 px-4 text-left text-sm font-bold text-gray-600 uppercase tracking-wider font-serif-khmer">ថ្ងៃនាំចេញ</th>
                                        <th className="py-3 px-4 text-left text-sm font-bold text-gray-600 uppercase tracking-wider">Description</th>
                                        <th className="py-3 px-4 text-left text-sm font-bold text-gray-600 uppercase tracking-wider font-serif-khmer">ចំនួនសរុប</th>
                                        <th className="py-3 px-4 text-left text-sm font-bold text-gray-600 uppercase tracking-wider font-serif-khmer">ស្ថានភាព</th>
                                    </tr>
                                </thead>
                                <tbody className="divide-y divide-gray-200">
                                    {filteredReports.length > 0 ? filteredReports.map(report => (
                                        <tr 
                                            key={report.id}
                                            onClick={() => handleSelectReport(report)}
                                            className={`cursor-pointer transition-colors duration-200 ${selectedReportId === report.id ? 'bg-indigo-100' : 'hover:bg-gray-50'}`}
                                        >
                                            <td className="py-3 px-4 whitespace-nowrap">{report.requesterName}</td>
                                            <td className="py-3 px-4 whitespace-nowrap">{report.campus}</td>
                                            <td className="py-3 px-4 whitespace-nowrap">{report.importDate}</td>
                                            <td className="py-3 px-4 whitespace-nowrap">{report.exportDate}</td>
                                            <td className="py-3 px-4 whitespace-nowrap">{formatItemsForDisplay(report.items)}</td>
                                            <td className="py-3 px-4 whitespace-nowrap text-center">{calculateTotalItems(report.items)}</td>
                                            <td className="py-3 px-4 whitespace-nowrap">
                                                <span className={`px-3 py-1 inline-flex text-xs leading-5 font-semibold rounded-full ${
                                                    (report.status || 'Process') === 'Done' ? 'bg-green-100 text-green-800' : 'bg-yellow-100 text-yellow-800'
                                                }`}>
                                                    {report.status || 'Process'}
                                                </span>
                                            </td>
                                        </tr>
                                    )) : (
                                        <tr>
                                            <td colSpan={7} className="text-center py-8 text-gray-500">
                                                {reports.length > 0 ? 'No matching reports found.' : 'No reports yet.'}
                                            </td>
                                        </tr>
                                    )}
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
            </div>
             <ConfirmationModal
                isOpen={isConfirmingDelete}
                onConfirm={handleConfirmDelete}
                onCancel={() => setIsConfirmingDelete(false)}
                title="Confirm Report Deletion"
                confirmButtonText="Delete"
            >
                {reportToDelete ? (
                    <>
                        <p>Are you sure you want to permanently delete the report for <strong className="text-indigo-600">{reportToDelete.requesterName}</strong> from campus <strong className="text-indigo-600">{reportToDelete.campus}</strong>?</p>
                        <p className="mt-4 text-sm text-gray-500">This action cannot be undone.</p>
                    </>
                ) : (
                    <p>Are you sure you want to delete this report? This action cannot be undone.</p>
                )}
            </ConfirmationModal>
            <ConfirmationModal
                isOpen={isConfirmingClearStock}
                onConfirm={handleConfirmClearStock}
                onCancel={() => setIsConfirmingClearStock(false)}
                title="Confirm Clear Stock"
                confirmButtonText="Clear All"
            >
                <p>Are you sure you want to permanently clear all stock data? This will set the quantity of all items to 0.</p>
                <p className="mt-4 text-sm text-gray-500">This action cannot be undone.</p>
            </ConfirmationModal>
        </>
    );
}
