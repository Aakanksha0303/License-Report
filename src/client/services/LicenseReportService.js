/**
 * Service for fetching license report data from ServiceNow
 * Replicates the exact data structure from Combined Report with additional sub-columns
 */
export class LicenseReportService {
    constructor() {
        this.baseURL = '/api/now/table';
      
        this.tableName = 'x_hete_hex_sn_lice_customer_licenses_track';
        
        // Exact filter from the original Combined Report
        this.reportFilter = 'valid_from<=javascript:gs.endOfToday()^valid_till>=javascript:gs.beginningOfToday()^license.renewal_date>=javascript:gs.beginningOfToday()^purchased_for=2a6d25541bd75510c8f821f72a4bcbd6^customer.active=true';
    }

    /**
     * Fetch the license report data with the exact same structure as Combined Report
     * Now includes additional fields for sub-columns
     */
    async fetchLicenseData() {
        try {
            const params = new URLSearchParams({
                sysparm_query: this.reportFilter,
                sysparm_fields: 'customer.name,customer.sys_id,license.license.name,license.license.sys_id,alloted_units,rpt_costxunits,cost_price_unit_v1,selling_price_unit,sys_id',
                sysparm_display_value: 'all',
                sysparm_limit: '1000'
            });

            const response = await fetch(`${this.baseURL}/${this.tableName}?${params}`, {
                method: 'GET',
                headers: {
                    'Content-Type': 'application/json',
                    'Accept': 'application/json',
                    'X-UserToken': window.g_ck || ''
                }
            });

            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }

            const data = await response.json();
            return this.processDataForPivotTable(data.result || []);
        } catch (error) {
            console.error('Error fetching license data:', error);
            throw error;
        }
    }

    /**
     * Process the raw data into pivot table format with sub-columns
     * Columns: Customer Names with sub-columns (Allocated Unit, Total Cost, Cost Price/Unit, Selling price/Unit)
     * Rows: License SKU Names  
     * Values: Aggregated values for each sub-column
     */
    processDataForPivotTable(rawData) {
        const customerMap = new Map();
        const licenseMap = new Map(); 
        const pivotData = new Map();

        // Process each record
        rawData.forEach(record => {
            const customerName = record['customer.name']?.display_value || 'Unknown Customer';
            const customerSysId = record['customer.sys_id']?.value;
            const licenseName = record['license.license.name']?.display_value || 'Unknown License';
            const licenseSysId = record['license.license.sys_id']?.value;
            
            // Parse all the sub-column values
            const allocatedUnits = parseFloat(record.alloted_units?.value || 0);
            const totalCost = parseFloat(record.rpt_costxunits?.value || 0);
            const costPriceUnit = parseFloat(record.cost_price_unit_v1?.value || 0);
            const sellingPriceUnit = parseFloat(record.selling_price_unit?.value || 0);

            // Store unique customers and licenses
            if (customerSysId) {
                customerMap.set(customerSysId, customerName);
            }
            if (licenseSysId) {
                licenseMap.set(licenseSysId, licenseName);
            }

            // Create pivot key
            const pivotKey = `${customerSysId}|${licenseSysId}`;
            
            // Aggregate values for same customer-license combination
            if (pivotData.has(pivotKey)) {
                const existing = pivotData.get(pivotKey);
                pivotData.set(pivotKey, {
                    allocatedUnits: existing.allocatedUnits + allocatedUnits,
                    totalCost: existing.totalCost + totalCost,
                    costPriceUnit: existing.costPriceUnit + costPriceUnit,
                    sellingPriceUnit: existing.sellingPriceUnit + sellingPriceUnit
                });
            } else {
                pivotData.set(pivotKey, {
                    allocatedUnits,
                    totalCost,
                    costPriceUnit,
                    sellingPriceUnit
                });
            }
        });

        // Convert to structured format
        const customers = Array.from(customerMap.entries()).map(([sysId, name]) => ({
            sysId,
            name
        })).sort((a, b) => a.name.localeCompare(b.name));

        const licenses = Array.from(licenseMap.entries()).map(([sysId, name]) => ({
            sysId,
            name
        })).sort((a, b) => a.name.localeCompare(b.name));

        // Build pivot table data structure
        const pivotTable = [];
        
        licenses.forEach(license => {
            const row = {
                license: license.name,
                licenseSysId: license.sysId,
                data: {},
                rowTotals: {
                    allocatedUnits: 0,
                    totalCost: 0,
                    costPriceUnit: 0,
                    sellingPriceUnit: 0
                }
            };

            customers.forEach(customer => {
                const pivotKey = `${customer.sysId}|${license.sysId}`;
                const values = pivotData.get(pivotKey) || {
                    allocatedUnits: 0,
                    totalCost: 0,
                    costPriceUnit: 0,
                    sellingPriceUnit: 0
                };
                
                row.data[customer.name] = values;
                
                // Add to row totals
                row.rowTotals.allocatedUnits += values.allocatedUnits;
                row.rowTotals.totalCost += values.totalCost;
                row.rowTotals.costPriceUnit += values.costPriceUnit;
                row.rowTotals.sellingPriceUnit += values.sellingPriceUnit;
            });

            // Only include rows that have data
            if (row.rowTotals.allocatedUnits > 0 || row.rowTotals.totalCost > 0) {
                pivotTable.push(row);
            }
        });

        // Calculate column totals
        const columnTotals = {};
        const grandTotals = {
            allocatedUnits: 0,
            totalCost: 0,
            costPriceUnit: 0,
            sellingPriceUnit: 0
        };

        customers.forEach(customer => {
            const customerTotals = {
                allocatedUnits: 0,
                totalCost: 0,
                costPriceUnit: 0,
                sellingPriceUnit: 0
            };
            
            pivotTable.forEach(row => {
                const values = row.data[customer.name] || {};
                customerTotals.allocatedUnits += values.allocatedUnits || 0;
                customerTotals.totalCost += values.totalCost || 0;
                customerTotals.costPriceUnit += values.costPriceUnit || 0;
                customerTotals.sellingPriceUnit += values.sellingPriceUnit || 0;
            });
            
            columnTotals[customer.name] = customerTotals;
            
            grandTotals.allocatedUnits += customerTotals.allocatedUnits;
            grandTotals.totalCost += customerTotals.totalCost;
            grandTotals.costPriceUnit += customerTotals.costPriceUnit;
            grandTotals.sellingPriceUnit += customerTotals.sellingPriceUnit;
        });

        return {
            customers: customers.map(c => c.name),
            licenses: licenses.map(l => l.name),
            pivotTable,
            columnTotals,
            grandTotals,
            rawDataCount: rawData.length
        };
    }

    /**
     * Get unique license SKUs from the report data
     */
    async getUniqueLicenses() {
        try {
            const data = await this.fetchLicenseData();
            return data.licenses;
        } catch (error) {
            console.error('Error fetching unique licenses:', error);
            throw error;
        }
    }

    /**
     * Get unique customers from the report data
     */
    async getUniqueCustomers() {
        try {
            const data = await this.fetchLicenseData();
            return data.customers;
        } catch (error) {
            console.error('Error fetching unique customers:', error);
            throw error;
        }
    }
}

export default LicenseReportService;