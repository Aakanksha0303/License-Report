import React, { useState, useEffect } from 'react';
import LicenseReportService from '../services/LicenseReportService';
import './PivotTable.css';

/**
 * Pivot Table Component that replicates the exact Combined Report structure with sub-columns
 * Shows Customer Names as columns with 4 sub-columns each, License SKUs as rows
 */
const PivotTable = () => {
    const [data, setData] = useState(null);
    const [loading, setLoading] = useState(true);
    const [error, setError] = useState(null);

    useEffect(() => {
        loadData();
    }, []);

    const loadData = async () => {
        try {
            setLoading(true);
            setError(null);
            
            const licenseService = new LicenseReportService();
            const pivotData = await licenseService.fetchLicenseData();
            
            setData(pivotData);
        } catch (err) {
            console.error('Error loading pivot data:', err);
            setError(err.message || 'Failed to load data');
        } finally {
            setLoading(false);
        }
    };

    const formatNumber = (value, isPrice = false) => {
        if (value === 0) return '';
        if (isPrice) {
            return value.toLocaleString(undefined, { 
                minimumFractionDigits: 2,
                maximumFractionDigits: 2 
            });
        }
        return value.toLocaleString(undefined, { 
            minimumFractionDigits: 0,
            maximumFractionDigits: 2 
        });
    };

    const exportToExcel = async () => {
        if (!data || !data.pivotTable) {
            alert('No data available to export');
            return;
        }

        try {
            // Dynamically import XLSX to avoid loading issues
            const XLSX = await import('xlsx');

            // Sub-column definitions
            const subColumns = [
                { key: 'allocatedUnits', label: 'Allocated Units' },
                { key: 'totalCost', label: 'Total Cost' },
                { key: 'costPriceUnit', label: 'Cost Price/Unit' },
                { key: 'sellingPriceUnit', label: 'Selling Price/Unit' }
            ];

            // Create Excel data structure
            const excelData = [];

            // Header Row 1: Customer Names (spanning 4 columns each)
            const headerRow1 = ['License SKU'];
            data.customers.forEach(customer => {
                headerRow1.push(customer, '', '', ''); // Customer name spans 4 columns
            });
            headerRow1.push('Total', '', '', ''); // Total spans 4 columns
            excelData.push(headerRow1);

            // Header Row 2: Sub-column labels
            const headerRow2 = [''];
            data.customers.forEach(() => {
                subColumns.forEach(subCol => {
                    headerRow2.push(subCol.label);
                });
            });
            subColumns.forEach(subCol => {
                headerRow2.push(subCol.label);
            });
            excelData.push(headerRow2);

            // Data rows
            data.pivotTable.forEach(row => {
                const dataRow = [row.license];
                
                // Customer data
                data.customers.forEach(customer => {
                    const customerData = row.data[customer] || {};
                    subColumns.forEach(subCol => {
                        const value = customerData[subCol.key] || 0;
                        dataRow.push(value === 0 ? '' : value);
                    });
                });

                // Row totals
                subColumns.forEach(subCol => {
                    const value = row.rowTotals[subCol.key] || 0;
                    dataRow.push(value === 0 ? '' : value);
                });

                excelData.push(dataRow);
            });

            // Total row
            const totalRow = ['Total'];
            data.customers.forEach(customer => {
                const customerTotals = data.columnTotals[customer] || {};
                subColumns.forEach(subCol => {
                    const value = customerTotals[subCol.key] || 0;
                    totalRow.push(value === 0 ? '' : value);
                });
            });
            subColumns.forEach(subCol => {
                const value = data.grandTotals[subCol.key] || 0;
                totalRow.push(value === 0 ? '' : value);
            });
            excelData.push(totalRow);

            // Create workbook and worksheet
            const wb = XLSX.utils.book_new();
            const ws = XLSX.utils.aoa_to_sheet(excelData);

            // Set column widths
            const colWidths = [{ wch: 50 }]; // License SKU column
            data.customers.forEach(() => {
                colWidths.push({ wch: 15 }, { wch: 15 }, { wch: 15 }, { wch: 15 }); // 4 sub-columns per customer
            });
            colWidths.push({ wch: 15 }, { wch: 15 }, { wch: 15 }, { wch: 15 }); // Total columns
            ws['!cols'] = colWidths;

            // Merge cells for customer headers
            const merges = [];
            let colIndex = 1;
            data.customers.forEach((customer, customerIndex) => {
                // Merge customer name across 4 columns
                merges.push({
                    s: { r: 0, c: colIndex },
                    e: { r: 0, c: colIndex + 3 }
                });
                colIndex += 4;
            });
            // Merge "Total" header across 4 columns
            merges.push({
                s: { r: 0, c: colIndex },
                e: { r: 0, c: colIndex + 3 }
            });
            ws['!merges'] = merges;

            // Add worksheet to workbook
            XLSX.utils.book_append_sheet(wb, ws, 'Combined Report');

            // Generate filename with current date
            const today = new Date();
            const dateStr = today.toISOString().split('T')[0]; // YYYY-MM-DD format
            const filename = `Combined_Report_${dateStr}.xlsx`;

            // Save file
            XLSX.writeFile(wb, filename);

            alert('Excel file downloaded successfully!');

        } catch (error) {
            console.error('Error exporting to Excel:', error);
            alert('Failed to export to Excel. Please try again.');
        }
    };

    const exportToPDF = async () => {
        if (!data || !data.pivotTable) {
            alert('No data available to export');
            return;
        }

        try {
            console.log('Starting multi-page PDF export...');

            // Create a new window for multi-page PDF
            const printWindow = window.open('', '_blank');
            
            if (!printWindow) {
                alert('Please allow pop-ups for PDF export to work');
                return;
            }

            const today = new Date();
            const dateStr = today.toLocaleDateString();

            // Sub-column definitions
            const subColumns = [
                { key: 'allocatedUnits', label: 'Allocated Units' },
                { key: 'totalCost', label: 'Total Cost' },
                { key: 'costPriceUnit', label: 'Cost Price/Unit' },
                { key: 'sellingPriceUnit', label: 'Selling Price/Unit' }
            ];

            // Split customers into groups of 3 for readable pages
            const customersPerPage = 3;
            const customerGroups = [];
            for (let i = 0; i < data.customers.length; i += customersPerPage) {
                customerGroups.push(data.customers.slice(i, i + customersPerPage));
            }

            // Generate HTML for all pages
            let allPagesHTML = '';

            customerGroups.forEach((customerGroup, pageIndex) => {
                const isLastPage = pageIndex === customerGroups.length - 1;
                
                // Create table for this page
                let pageHTML = `
                    <div class="page" ${isLastPage ? '' : 'style="page-break-after: always;"'}>
                        <div class="header">
                            <h1>Combined Report - Page ${pageIndex + 1} of ${customerGroups.length}</h1>
                            <p>License Allocation Report</p>
                            <p>Generated on: ${dateStr}</p>
                            <p>Customers: ${customerGroup.join(', ')}</p>
                        </div>
                        
                        <table class="pivot-table">
                            <thead>
                                <!-- Customer header row -->
                                <tr>
                                    <th class="row-header" rowspan="2">License SKU</th>
                `;

                // Add customer headers for this page
                customerGroup.forEach(customer => {
                    pageHTML += `<th class="customer-header" colspan="4">${customer}</th>`;
                });
                
                pageHTML += `
                                </tr>
                                <!-- Sub-column header row -->
                                <tr>
                `;

                // Add sub-column headers for this page
                customerGroup.forEach(() => {
                    subColumns.forEach(subCol => {
                        pageHTML += `<th class="sub-column-header">${subCol.label}</th>`;
                    });
                });
                
                pageHTML += `
                                </tr>
                            </thead>
                            <tbody>
                `;

                // Add data rows for this page
                data.pivotTable.forEach(row => {
                    pageHTML += `
                        <tr class="data-row">
                            <td class="row-label">${row.license}</td>
                    `;

                    customerGroup.forEach(customer => {
                        const customerData = row.data[customer] || {};
                        subColumns.forEach(subCol => {
                            const value = customerData[subCol.key] || 0;
                            const formattedValue = value === 0 ? '' : (subCol.key.includes('price') || subCol.key === 'totalCost' ? 
                                value.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }) :
                                value.toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 2 }));
                            pageHTML += `<td class="data-cell">${formattedValue}</td>`;
                        });
                    });

                    pageHTML += `</tr>`;
                });

                // Add totals row for this page
                pageHTML += `
                            <tr class="totals-row">
                                <td class="total-label"><strong>Total</strong></td>
                `;

                customerGroup.forEach(customer => {
                    const customerTotals = data.columnTotals[customer] || {};
                    subColumns.forEach(subCol => {
                        const value = customerTotals[subCol.key] || 0;
                        const formattedValue = value === 0 ? '' : (subCol.key.includes('price') || subCol.key === 'totalCost' ? 
                            value.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }) :
                            value.toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 2 }));
                        pageHTML += `<td class="column-total"><strong>${formattedValue}</strong></td>`;
                    });
                });

                pageHTML += `
                            </tr>
                            </tbody>
                        </table>
                    </div>
                `;

                allPagesHTML += pageHTML;
            });

            // Create complete printable HTML
            const printHTML = `
                <!DOCTYPE html>
                <html>
                <head>
                    <title>Combined Report - ${dateStr}</title>
                    <style>
                        * {
                            box-sizing: border-box;
                        }
                        body { 
                            font-family: Arial, sans-serif; 
                            margin: 0;
                            padding: 20px;
                            background: white;
                            font-size: 12px;
                        }
                        .page {
                            width: 100%;
                            min-height: 100vh;
                            padding: 20px 0;
                        }
                        .header { 
                            text-align: center; 
                            margin-bottom: 20px; 
                            border-bottom: 2px solid #333;
                            padding-bottom: 10px;
                        }
                        .header h1 { 
                            margin: 0; 
                            color: #333; 
                            font-size: 20px;
                        }
                        .header p { 
                            margin: 5px 0; 
                            color: #666; 
                            font-size: 12px;
                        }
                        .pivot-table {
                            width: 100%;
                            border-collapse: collapse;
                            font-size: 11px;
                            background-color: #ffffff;
                            margin-bottom: 30px;
                        }
                        .pivot-table th, .pivot-table td {
                            border: 1px solid #dee2e6;
                            padding: 6px 8px;
                            text-align: center;
                            word-wrap: break-word;
                        }
                        .row-header, .row-label {
                            text-align: left !important;
                            font-weight: 500;
                            background-color: #f8f9fa;
                            width: 40%;
                            font-size: 10px;
                            line-height: 1.3;
                        }
                        .customer-header {
                            background-color: #d1ecf1 !important;
                            color: #0c5460 !important;
                            font-weight: 700;
                            font-size: 12px;
                        }
                        .sub-column-header {
                            background-color: #f8f9fa !important;
                            font-weight: 600;
                            color: #495057;
                            font-size: 9px;
                            width: 15%;
                        }
                        .data-cell {
                            text-align: right !important;
                            font-family: 'Courier New', monospace;
                            font-size: 10px;
                        }
                        .column-total {
                            text-align: right !important;
                            font-weight: 700;
                            color: #155724;
                            background-color: #d4edda !important;
                            font-size: 10px;
                        }
                        .total-label {
                            background-color: #c3e6cb !important;
                            font-weight: 700;
                            color: #155724;
                            text-align: left !important;
                            font-size: 11px;
                        }
                        .totals-row {
                            background-color: #d4edda;
                        }
                        @page {
                            size: A4 landscape;
                            margin: 0.75in;
                        }
                        @media print {
                            body { 
                                padding: 0;
                                font-size: 10px;
                            }
                            .page {
                                padding: 10px 0;
                            }
                            .pivot-table { 
                                font-size: 9px;
                            }
                            .pivot-table th, .pivot-table td {
                                padding: 4px 6px;
                            }
                        }
                    </style>
                </head>
                <body>
                    ${allPagesHTML}
                    <script>
                        window.onload = function() {
                            setTimeout(function() {
                                window.print();
                                setTimeout(function() {
                                    window.close();
                                }, 1000);
                            }, 500);
                        };
                    </script>
                </body>
                </html>
            `;

            printWindow.document.write(printHTML);
            printWindow.document.close();

            alert(`PDF will be generated in ${customerGroups.length} readable pages (${customersPerPage} customers per page). Each page will be clearly readable with proper font sizes.`);

        } catch (error) {
            console.error('Error exporting to PDF:', error);
            alert(`Failed to export to PDF: ${error.message}. Please try using your browser's print function instead.`);
        }
    };

    if (loading) {
        return (
            <div className="pivot-container">
                <div className="loading-state">
                    <div className="loading-spinner"></div>
                    <p>Loading Combined Report data...</p>
                </div>
            </div>
        );
    }

    if (error) {
        return (
            <div className="pivot-container">
                <div className="error-state">
                    <h3>Error Loading Report</h3>
                    <p>{error}</p>
                    <button onClick={loadData} className="retry-button">
                        Retry
                    </button>
                </div>
            </div>
        );
    }

    if (!data || !data.pivotTable || data.pivotTable.length === 0) {
        return (
            <div className="pivot-container">
                <div className="no-data-state">
                    <h3>No Data Available</h3>
                    <p>No license allocation data found matching the report criteria.</p>
                    <button onClick={loadData} className="retry-button">
                        Refresh
                    </button>
                </div>
            </div>
        );
    }

    // Sub-column definitions
    const subColumns = [
        { key: 'allocatedUnits', label: 'Allocated Units', isPrice: false },
        { key: 'totalCost', label: 'Total Cost', isPrice: true },
        { key: 'costPriceUnit', label: 'Cost Price/Unit', isPrice: true },
        { key: 'sellingPriceUnit', label: 'Selling Price/Unit', isPrice: true }
    ];

    return (
        <div className="pivot-container">
            {/* Report Header with Export Buttons */}
            <div style={{ 
                display: 'flex', 
                justifyContent: 'space-between', 
                alignItems: 'center', 
                backgroundColor: 'white', 
                padding: '20px', 
                border: '1px solid #ccc', 
                borderRadius: '5px',
                marginBottom: '20px'
            }}>
                <h1 style={{ margin: '0', color: '#333' }}>Combined Report</h1>
                <div style={{ display: 'flex', gap: '10px' }}>
                    <button 
                        onClick={exportToExcel} 
                        style={{
                            backgroundColor: '#28a745',
                            color: 'white',
                            border: 'none',
                            padding: '12px 20px',
                            borderRadius: '5px',
                            cursor: 'pointer',
                            fontSize: '14px',
                            fontWeight: '600'
                        }}
                    >
                        📊 Export to Excel
                    </button>
                    <button 
                        onClick={exportToPDF} 
                        style={{
                            backgroundColor: '#dc3545',
                            color: 'white',
                            border: 'none',
                            padding: '12px 20px',
                            borderRadius: '5px',
                            cursor: 'pointer',
                            fontSize: '14px',
                            fontWeight: '600'
                        }}
                    >
                        📄 Export to PDF
                    </button>
                </div>
            </div>

            {/* Pivot Table */}
            <div className="pivot-table-wrapper">
                <table className="pivot-table">
                    <thead>
                        {/* Customer header row */}
                        <tr>
                            <th className="row-header sticky-header" rowSpan="2">License SKU</th>
                            {data.customers.map((customer, index) => (
                                <th key={index} className="customer-header sticky-header" colSpan="4">
                                    {customer}
                                </th>
                            ))}
                            <th className="total-header sticky-header" colSpan="4">Total</th>
                        </tr>
                        {/* Sub-column header row */}
                        <tr>
                            {data.customers.map((customer, customerIndex) => (
                                subColumns.map((subCol, subIndex) => (
                                    <th key={`${customerIndex}-${subIndex}`} className="sub-column-header sticky-header">
                                        {subCol.label}
                                    </th>
                                ))
                            ))}
                            {subColumns.map((subCol, subIndex) => (
                                <th key={`total-${subIndex}`} className="sub-total-header sticky-header">
                                    {subCol.label}
                                </th>
                            ))}
                        </tr>
                    </thead>
                    <tbody>
                        {data.pivotTable.map((row, rowIndex) => (
                            <tr key={rowIndex} className="data-row">
                                <td className="row-label sticky-column">
                                    {row.license}
                                </td>
                                {data.customers.map((customer, customerIndex) => {
                                    const customerData = row.data[customer] || {};
                                    return subColumns.map((subCol, subIndex) => (
                                        <td key={`${customerIndex}-${subIndex}`} className="data-cell">
                                            {formatNumber(customerData[subCol.key] || 0, subCol.isPrice)}
                                        </td>
                                    ));
                                })}
                                {subColumns.map((subCol, subIndex) => (
                                    <td key={`row-total-${subIndex}`} className="row-total">
                                        {formatNumber(row.rowTotals[subCol.key] || 0, subCol.isPrice)}
                                    </td>
                                ))}
                            </tr>
                        ))}
                    </tbody>
                    <tfoot>
                        <tr className="totals-row">
                            <td className="total-label sticky-column">
                                <strong>Total</strong>
                            </td>
                            {data.customers.map((customer, customerIndex) => {
                                const customerTotals = data.columnTotals[customer] || {};
                                return subColumns.map((subCol, subIndex) => (
                                    <td key={`col-total-${customerIndex}-${subIndex}`} className="column-total">
                                        <strong>{formatNumber(customerTotals[subCol.key] || 0, subCol.isPrice)}</strong>
                                    </td>
                                ));
                            })}
                            {subColumns.map((subCol, subIndex) => (
                                <td key={`grand-total-${subIndex}`} className="grand-total">
                                    <strong>{formatNumber(data.grandTotals[subCol.key] || 0, subCol.isPrice)}</strong>
                                </td>
                            ))}
                        </tr>
                    </tfoot>
                </table>
            </div>

            {/* Report Summary */}
            <div className="report-summary">
                <div className="summary-stats">
                    <div className="stat-item">
                        <span className="stat-label">Unique Customers:</span>
                        <span className="stat-value">{data.customers.length}</span>
                    </div>
                    <div className="stat-item">
                        <span className="stat-label">Unique License SKUs:</span>
                        <span className="stat-value">{data.licenses.length}</span>
                    </div>
                    <div className="stat-item">
                        <span className="stat-label">Total Allocated Units:</span>
                        <span className="stat-value">{formatNumber(data.grandTotals.allocatedUnits)}</span>
                    </div>
                    <div className="stat-item">
                        <span className="stat-label">Total Cost:</span>
                        <span className="stat-value">{formatNumber(data.grandTotals.totalCost, true)}</span>
                    </div>
                </div>
            </div>

            {/* Refresh Button */}
            <div className="report-actions">
                <button onClick={loadData} className="refresh-button">
                    Refresh Data
                </button>
            </div>
        </div>
    );
};

export default PivotTable;