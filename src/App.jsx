import React, { useState, useEffect, useMemo } from 'react';
import {
  Package,
  Plus,
  Trash2,
  Calculator,
  Scale,
  Settings,
  RefreshCcw,
  Download,
  AlertTriangle,
  X,
  Info,
  Upload,
  FileSpreadsheet,
  Edit2
} from 'lucide-react';
import * as XLSX from 'xlsx';


const App = () => {
  // --- State ---
  const [globalSettings, setGlobalSettings] = useState({
    cnToCnCost: 0,   // Total RMB for the box (Fixed)
    cnToBdRate: 0,   // Local (BDT) per KG
    currencyRate: 1, // RMB to Local conversion
    otherCost: 0,    // Local (BDT)
    totalBoxWeight: 0,
  });

  const [products, setProducts] = useState([]);
  const [newProduct, setNewProduct] = useState({
    sn: '',
    name: '',
    weight: '',
    weightUnit: 'gram', // 'kg' or 'gram' (Default: gram)
    cost: '',
    quantity: ''
  });

  const [editingId, setEditingId] = useState(null); // Track which product is being edited
  const [showOverweightPopup, setShowOverweightPopup] = useState(false);
  const [showExportMenu, setShowExportMenu] = useState(false);

  // --- Persistence ---
  useEffect(() => {
    const savedData = localStorage.getItem('shipmentCalculatorData_v3'); // New key for v3 data structure
    if (savedData) {
      try {
        const parsed = JSON.parse(savedData);
        // Migration: map old cnToCnRate to new cnToCnCost if needed, or just default to 0
        setGlobalSettings(prev => ({
          ...prev,
          ...parsed.globalSettings,
          cnToCnCost: parsed.globalSettings.cnToCnCost ?? parsed.globalSettings.cnToCnRate ?? 0
        }));
        setProducts(parsed.products || []);
      } catch (e) {
        console.error("Failed to load saved data");
      }
    }
  }, []);

  useEffect(() => {
    localStorage.setItem('shipmentCalculatorData_v3', JSON.stringify({ globalSettings, products }));
  }, [globalSettings, products]);

  // --- Handlers ---
  const handleGlobalChange = (e) => {
    const { name, value } = e.target;
    setGlobalSettings(prev => ({
      ...prev,
      [name]: parseFloat(value) || 0
    }));
  };

  const handleNewProductChange = (e) => {
    const { name, value } = e.target;
    setNewProduct(prev => ({
      ...prev,
      [name]: value
    }));
  };

  const addProduct = (e) => {
    e.preventDefault();
    if (!newProduct.name || !newProduct.weight || !newProduct.cost || !newProduct.quantity) return;

    let weightInKg = parseFloat(newProduct.weight);
    if (newProduct.weightUnit === 'gram') {
      weightInKg = weightInKg / 1000;
    }

    const payload = {
      sn: newProduct.sn,
      name: newProduct.name,
      weight: weightInKg, // Store in KG
      cost: parseFloat(newProduct.cost),
      quantity: parseInt(newProduct.quantity)
    };

    if (editingId) {
      // Update existing
      setProducts(products.map(p => p.id === editingId ? { ...p, ...payload } : p));
      setEditingId(null);
    } else {
      // Create new
      const product = {
        id: crypto.randomUUID(),
        sn: payload.sn || (products.length + 1).toString(),
        ...payload
      };
      setProducts([...products, product]);
    }

    setNewProduct({ sn: '', name: '', weight: '', weightUnit: 'gram', cost: '', quantity: '' });
  };

  const startEdit = (product) => {
    setEditingId(product.id);
    let displayWeight = product.weight;
    let displayUnit = 'kg';

    // Simple heuristic: if weight is small (< 1kg) and likely intended as grams, show as grams
    // Or just strictly convert if it was originally grams? 
    // For now, let's just keep it simple: always load as KG unless user changes it, 
    // OR if we want to be fancy, check if it's a clean gram number.
    // Let's stick to standard KG for edit to avoid confusion, or convert to grams if < 1.
    if (product.weight < 1) {
      displayWeight = product.weight * 1000;
      displayUnit = 'gram';
    }

    setNewProduct({
      sn: product.sn,
      name: product.name,
      weight: displayWeight.toString(),
      weightUnit: displayUnit,
      cost: product.cost.toString(),
      quantity: product.quantity.toString()
    });

    // Scroll to form
    document.querySelector('form')?.scrollIntoView({ behavior: 'smooth' });
  };

  const cancelEdit = () => {
    setEditingId(null);
    setNewProduct({ sn: '', name: '', weight: '', weightUnit: 'gram', cost: '', quantity: '' });
  };

  const handleImport = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsName = wb.SheetNames[0];
        const ws = wb.Sheets[wsName];
        const data = XLSX.utils.sheet_to_json(ws);

        const newProducts = data.map((row, idx) => {
          let weight = parseFloat(row.Weight || 0);
          const unit = row.Unit ? row.Unit.toLowerCase() : 'kg';
          if (unit === 'g' || unit === 'gram' || unit === 'grams') {
            weight = weight / 1000;
          }

          return {
            id: crypto.randomUUID(),
            sn: row.SKU ? row.SKU.toString() : (products.length + idx + 1).toString(),
            name: row.Name || 'Unknown Product',
            weight: weight,
            cost: parseFloat(row.Cost || 0),
            quantity: parseInt(row.Quantity || 1)
          };
        }).filter(p => p.name && p.weight > 0);

        setProducts(prev => [...prev, ...newProducts]);
        alert(`Successfully imported ${newProducts.length} products.`);
      } catch (err) {
        console.error(err);
        alert("Failed to parse file. Please ensure it matches the template.");
      }
    };
    reader.readAsBinaryString(file);
    e.target.value = null; // Reset input
  };

  const downloadTemplate = () => {
    const ws = XLSX.utils.json_to_sheet([
      { SKU: '001', Name: 'Sample Product', Weight: 0.5, Unit: 'kg', Cost: 50, Quantity: 10 },
      { SKU: '002', Name: 'Small Item', Weight: 500, Unit: 'gram', Cost: 10, Quantity: 5 }
    ]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Template");
    XLSX.writeFile(wb, "Freight_Template.xlsx");
  };

  const removeProduct = (id) => {
    setProducts(products.filter(p => p.id !== id));
    if (editingId === id) cancelEdit();
  };

  const clearAll = () => {
    if (confirm("Clear all products?")) {
      setProducts([]);
    }
  };

  // --- Calculations ---
  const calculations = useMemo(() => {
    // 1. Total Net Weight of all products
    const totalNetWeight = products.reduce((sum, p) => sum + (p.weight * p.quantity), 0);

    // 2. Extra Weight (Tare)
    const rawExtraWeight = globalSettings.totalBoxWeight - totalNetWeight;
    const extraWeight = Math.max(0, rawExtraWeight);

    // 3. Weight Distribution Factor
    const distributionFactor = totalNetWeight > 0 ? (extraWeight / totalNetWeight) : 0;

    // Process each product
    const processedProducts = products.map(product => {
      // Net Weight (Single Unit)
      const unitNetWeight = product.weight;

      // Allocated Extra Weight (Tare Share) per Unit
      const unitTareWeight = unitNetWeight * distributionFactor;

      // Per Product Box Weight (Gross Weight per Unit)
      const unitGrossWeight = unitNetWeight + unitTareWeight;

      // Base Cost in Local Currency (converted from RMB)
      const unitBaseCostLocal = product.cost * globalSettings.currencyRate;

      // --- Shipping Costs Calculation (2 Parts) ---

      // 1. Supplier -> CN Wirehouse (Fixed Box Cost Distributed by Weight)
      // Share = (Unit Gross Weight / Total Box Weight) * Total Box Cost
      const unitWeightShare = globalSettings.totalBoxWeight > 0
        ? unitGrossWeight / globalSettings.totalBoxWeight
        : 0;

      const costCnToCn = (unitWeightShare * globalSettings.cnToCnCost) * globalSettings.currencyRate;

      // 2. CN -> BD (Per KG Rate)
      const costCnToBd = unitGrossWeight * globalSettings.cnToBdRate;

      // Total Shipping for this unit
      const totalShippingCost = costCnToCn + costCnToBd;

      // --- Other Costs ---
      // Distributed based on weight share
      const allocatedOtherCost = unitWeightShare * globalSettings.otherCost;

      // Per Product Extra Cost (Shipping + Other)
      const unitExtraCost = totalShippingCost + allocatedOtherCost;

      // Per Product Gross Cost
      const unitGrossCost = unitBaseCostLocal + unitExtraCost;

      // Total Product Cost (Line Total)
      const totalLineCost = unitGrossCost * product.quantity;
      const totalRmbParams = product.cost * product.quantity;

      return {
        ...product,
        unitTareWeight,
        unitGrossWeight,
        unitBaseCostLocal,
        costCnToCn,
        costCnToBd,
        totalShippingCost,
        allocatedOtherCost,
        unitExtraCost,
        unitGrossCost,
        totalLineCost,
        totalRmbParams
      };
    });

    // Totals for Summary
    const totalGrossCost = processedProducts.reduce((sum, p) => sum + p.totalLineCost, 0);
    const totalShippingCost = processedProducts.reduce((sum, p) => sum + (p.totalShippingCost * p.quantity), 0);
    const totalOtherCost = processedProducts.reduce((sum, p) => sum + (p.allocatedOtherCost * p.quantity), 0);
    const totalRmbParams = processedProducts.reduce((sum, p) => sum + p.totalRmbParams, 0);

    return {
      totalNetWeight,
      extraWeight,
      processedProducts,
      totalGrossCost,
      totalShippingCost,
      totalOtherCost,
      totalRmbParams
    };
  }, [products, globalSettings]);

  // --- Effects ---
  useEffect(() => {
    if (globalSettings.totalBoxWeight > 0 && calculations.totalNetWeight > globalSettings.totalBoxWeight) {
      setShowOverweightPopup(true);
    } else {
      setShowOverweightPopup(false);
    }
  }, [calculations.totalNetWeight, globalSettings.totalBoxWeight]);


  // --- Export Functions ---
  const getExportData = () => {
    return calculations.processedProducts.map(p => ({
      Product: p.name,
      Qty: p.quantity,
      'Unit Price (RMB)': parseFloat(p.cost.toFixed(2)),
      'Total Price (RMB)': parseFloat(p.totalRmbParams.toFixed(2)),
      SKU: p.sn,
      'Gross Weight (KG)': parseFloat(p.unitGrossWeight.toFixed(3)),
      'Shipment Cost (BDT)': parseFloat(p.totalShippingCost.toFixed(2)),
      'Extra Cost (BDT)': parseFloat(p.allocatedOtherCost.toFixed(2)),
      'Total Cost (BDT)': parseFloat(p.totalLineCost.toFixed(2)),
      'Per Unit Price (BDT)': parseFloat(p.unitGrossCost.toFixed(2))
    }));
  };

  const exportToExcel = () => {
    const data = getExportData();
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Invoice");
    XLSX.writeFile(wb, "Freight_Invoice.xlsx");
    setShowExportMenu(false);
  };

  const exportToCSV = () => {
    const data = getExportData();
    const ws = XLSX.utils.json_to_sheet(data);
    const csv = XLSX.utils.sheet_to_csv(ws);
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement("a");
    const url = URL.createObjectURL(blob);
    link.setAttribute("href", url);
    link.setAttribute("download", "Freight_Invoice.csv");
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    setShowExportMenu(false);
  };

  // Format currency helper
  const fmt = (num) => num ? num.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : '0.00';
  const fmtW = (num) => num ? num.toLocaleString('en-US', { minimumFractionDigits: 3, maximumFractionDigits: 3 }) : '0.000';

  return (
    <div className="min-h-screen bg-slate-50 text-slate-800 font-sans selection:bg-blue-100 relative">

      {/* Overweight Popup */}
      {showOverweightPopup && (
        <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center p-4 backdrop-blur-sm">
          <div className="bg-white rounded-2xl shadow-xl max-w-md w-full overflow-hidden animate-in fade-in zoom-in duration-200">
            <div className="bg-red-50 p-6 flex flex-col items-center text-center border-b border-red-100">
              <div className="w-12 h-12 bg-red-100 rounded-full flex items-center justify-center mb-4">
                <AlertTriangle className="w-6 h-6 text-red-600" />
              </div>
              <h3 className="text-lg font-bold text-red-900">Overweight Warning</h3>
              <p className="text-sm text-red-700 mt-2">
                Total product weight ({fmtW(calculations.totalNetWeight)} kg) exceeds the Box Weight ({fmtW(globalSettings.totalBoxWeight)} kg).
              </p>
            </div>
            <div className="p-4 bg-white flex justify-end">
              <button
                onClick={() => setShowOverweightPopup(false)}
                className="px-4 py-2 bg-slate-900 text-white text-sm font-medium rounded-lg hover:bg-slate-800 transition-colors"
              >
                Dismiss
              </button>
            </div>
          </div>
        </div>
      )}

      <div className="max-w-7xl mx-auto p-4 md:p-8">

        {/* Header */}
        <header className="mb-8 flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div>
            <h1 className="text-3xl font-bold text-slate-900 flex items-center gap-3">
              <Package className="w-8 h-8 text-blue-600" />
              Freight Calculator
            </h1>
            <p className="text-slate-500 mt-1">International shipment cost distribution & pricing.</p>
          </div>
          <div className="flex items-center gap-2">
            <div className="relative">
              <input
                type="file"
                id="import-file"
                accept=".xlsx, .xls, .csv"
                className="hidden"
                onChange={handleImport}
              />
              <button
                onClick={() => document.getElementById('import-file').click()}
                className="flex items-center gap-2 text-sm text-blue-600 hover:bg-blue-50 px-3 py-2 rounded-lg transition-colors border border-blue-100"
              >
                <Upload className="w-4 h-4" /> Import
              </button>
            </div>
            <button
              onClick={downloadTemplate}
              title="Download Import Template"
              className="text-slate-400 hover:text-blue-600 p-2 rounded-lg transition-colors"
            >
              <FileSpreadsheet className="w-5 h-5" />
            </button>
            <button
              onClick={clearAll}
              className="flex items-center gap-2 text-sm text-red-600 hover:bg-red-50 px-3 py-2 rounded-lg transition-colors"
            >
              <RefreshCcw className="w-4 h-4" /> Reset
            </button>
          </div>
        </header>

        {/* Main Grid */}
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-8">

          {/* Left Column: Settings & Input (4 cols) */}
          <div className="lg:col-span-4 space-y-6">

            {/* Global Settings Card */}
            <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
              <div className="bg-slate-100 px-6 py-4 border-b border-slate-200 flex items-center gap-2">
                <Settings className="w-5 h-5 text-slate-600" />
                <h2 className="font-semibold text-slate-700">Shipment Settings</h2>
              </div>
              <div className="p-6 space-y-4">

                {/* Currency Section */}
                <div className="pb-4 border-b border-slate-100">
                  <label className="block text-xs font-semibold uppercase tracking-wider text-slate-500 mb-1">
                    Currency Rate (RMB → BDT)
                  </label>
                  <div className="relative">
                    <span className="absolute left-3 top-2.5 text-slate-400 font-bold">×</span>
                    <input
                      type="number"
                      name="currencyRate"
                      value={globalSettings.currencyRate || ''}
                      onChange={handleGlobalChange}
                      placeholder="1.00"
                      className="w-full pl-8 pr-4 py-2 border border-slate-200 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none transition-all"
                    />
                  </div>
                </div>

                {/* Shipping Rates Group */}
                <div className="space-y-3 pb-4 border-b border-slate-100">
                  <label className="block text-xs font-semibold uppercase tracking-wider text-slate-500">
                    Shipping Rates
                  </label>

                  <div className="grid grid-cols-1 gap-3">
                    <div className="relative">
                      <label className="text-[10px] text-slate-400 font-medium mb-1 block">
                        Supplier -&gt; CN Wirehouse (Total Box Cost)
                      </label>
                      <input
                        type="number"
                        name="cnToCnCost"
                        value={globalSettings.cnToCnCost || ''}
                        onChange={handleGlobalChange}
                        placeholder="Total RMB"
                        className="w-full pl-4 pr-12 py-2 text-sm border border-slate-200 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none"
                      />
                      <span className="absolute right-3 top-8 text-xs font-bold text-slate-400">RMB</span>
                    </div>

                    <div className="relative">
                      <label className="text-[10px] text-slate-400 font-medium mb-1 block">
                        CN -&gt; BD (Per KG Rate)
                      </label>
                      <input
                        type="number"
                        name="cnToBdRate"
                        value={globalSettings.cnToBdRate || ''}
                        onChange={handleGlobalChange}
                        placeholder="Per KG BDT"
                        className="w-full pl-4 pr-12 py-2 text-sm border border-slate-200 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none"
                      />
                      <span className="absolute right-3 top-8 text-xs font-bold text-slate-400">BDT</span>
                    </div>
                  </div>
                </div>

                <div className="grid grid-cols-1 gap-4">
                  <div>
                    <label className="block text-xs font-semibold uppercase tracking-wider text-slate-500 mb-1">
                      Box Weight (KG)
                    </label>
                    <div className="relative">
                      <span className="absolute left-2 top-2.5 text-slate-400"><Scale className="w-4 h-4" /></span>
                      <input
                        type="number"
                        name="totalBoxWeight"
                        value={globalSettings.totalBoxWeight || ''}
                        onChange={handleGlobalChange}
                        placeholder="KG"
                        className="w-full pl-8 pr-2 py-2 border border-slate-200 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none"
                      />
                    </div>
                  </div>

                  <div>
                    <label className="block text-xs font-semibold uppercase tracking-wider text-slate-500 mb-1">
                      Others (including BD to delivery location)
                    </label>
                    <div className="relative">
                      <span className="absolute left-3 top-2.5 text-slate-400 font-bold text-xs">BDT</span>
                      <input
                        type="number"
                        name="otherCost"
                        value={globalSettings.otherCost || ''}
                        onChange={handleGlobalChange}
                        placeholder="0.00"
                        className="w-full pl-11 pr-4 py-2 border border-slate-200 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none"
                      />
                    </div>
                  </div>
                </div>
                {calculations.extraWeight > 0 && (
                  <p className="text-xs text-slate-400 font-medium text-center">
                    Net: {fmtW(calculations.totalNetWeight)} | Box Wgt: {fmtW(calculations.extraWeight)}
                  </p>
                )}
              </div>
            </div>

            {/* Add/Edit Product Form */}
            <form onSubmit={addProduct} className={`bg-white rounded-2xl shadow-sm border ${editingId ? 'border-amber-200 ring-4 ring-amber-50' : 'border-slate-200'} overflow-hidden transition-all duration-300`}>
              <div className={`${editingId ? 'bg-amber-500 border-amber-600' : 'bg-blue-600 border-blue-700'} px-6 py-4 border-b flex items-center justify-between`}>
                <div className="flex items-center gap-2">
                  {editingId ? <Edit2 className="w-5 h-5 text-white" /> : <Plus className="w-5 h-5 text-blue-100" />}
                  <h2 className="font-semibold text-white">{editingId ? 'Edit Product' : 'Add Product'}</h2>
                </div>
                {editingId && (
                  <button type="button" onClick={cancelEdit} className="text-white/80 hover:text-white">
                    <X className="w-5 h-5" />
                  </button>
                )}
              </div>
              <div className="p-6 space-y-4">
                <div className="grid grid-cols-2 gap-4">
                  <div>
                    <label className="block text-xs text-slate-500 mb-1">SKU (Opt)</label>
                    <input name="sn" value={newProduct.sn} onChange={handleNewProductChange} className="w-full px-3 py-2 border border-slate-200 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none" placeholder="001" />
                  </div>
                  <div>
                    <label className="block text-xs text-slate-500 mb-1">Quantity</label>
                    <input type="number" name="quantity" value={newProduct.quantity} onChange={handleNewProductChange} className="w-full px-3 py-2 border border-slate-200 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none" placeholder="1" required />
                  </div>
                </div>

                <div>
                  <label className="block text-xs text-slate-500 mb-1">Product Name</label>
                  <input name="name" value={newProduct.name} onChange={handleNewProductChange} className="w-full px-3 py-2 border border-slate-200 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none" placeholder="Item description" required />
                </div>

                <div className="grid grid-cols-2 gap-4">
                  <div>
                    <label className="block text-xs text-slate-500 mb-1">Weight</label>
                    <div className="flex">
                      <input type="number" step="0.001" name="weight" value={newProduct.weight} onChange={handleNewProductChange} className="w-full px-3 py-2 border border-slate-200 rounded-l-lg focus:ring-2 focus:ring-blue-500 outline-none" placeholder="0.5" required />
                      <select
                        name="weightUnit"
                        value={newProduct.weightUnit}
                        onChange={handleNewProductChange}
                        className="bg-slate-100 border border-l-0 border-slate-200 rounded-r-lg px-2 text-xs font-medium text-slate-600 focus:ring-2 focus:ring-blue-500 outline-none"
                      >
                        <option value="kg">KG</option>
                        <option value="gram">Gram</option>
                      </select>
                    </div>
                  </div>
                  <div>
                    <label className="block text-xs text-slate-500 mb-1">Cost (RMB)</label>
                    <input type="number" step="0.01" name="cost" value={newProduct.cost} onChange={handleNewProductChange} className="w-full px-3 py-2 border border-slate-200 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none" placeholder="10.00" required />
                  </div>
                </div>

                <div className="flex gap-2 mt-2">
                  {editingId && (
                    <button
                      type="button"
                      onClick={cancelEdit}
                      className="flex-1 bg-slate-100 hover:bg-slate-200 text-slate-600 font-semibold py-2.5 rounded-lg transition-colors"
                    >
                      Cancel
                    </button>
                  )}
                  <button
                    type="submit"
                    className={`flex-1 ${editingId ? 'bg-amber-500 hover:bg-amber-600' : 'bg-blue-600 hover:bg-blue-700'} text-white font-semibold py-2.5 rounded-lg transition-colors flex items-center justify-center gap-2`}
                  >
                    {editingId ? <Edit2 className="w-4 h-4" /> : <Plus className="w-4 h-4" />}
                    {editingId ? 'Update Product' : 'Add to Shipment'}
                  </button>
                </div>
              </div>
            </form>

            {/* Quick Summary Small (Mobile only mostly) */}
            <div className="bg-slate-900 text-slate-200 rounded-2xl p-6 lg:hidden">
              <div className="flex justify-between items-end mb-2">
                <span className="text-sm uppercase tracking-wider text-slate-400">Total BDT</span>
                <span className="text-3xl font-bold text-white">{fmt(calculations.totalGrossCost)}</span>
              </div>
            </div>

          </div>

          {/* Right Column: Results Table (8 cols) */}
          <div className="lg:col-span-8 flex flex-col gap-6">

            {/* Desktop Summary Cards */}
            <div className="hidden lg:grid grid-cols-3 gap-4">
              <div className="bg-white p-4 rounded-xl shadow-sm border border-slate-200">
                <p className="text-xs text-slate-500 uppercase font-bold tracking-wider mb-1">Total Shipping (BDT)</p>
                <p className="text-2xl font-bold text-slate-700">{fmt(calculations.totalShippingCost)}</p>
                <p className="text-[10px] text-slate-400 mt-1">Includes Supplier→CN, CN→BD</p>
              </div>
              <div className="bg-white p-4 rounded-xl shadow-sm border border-slate-200">
                <p className="text-xs text-slate-500 uppercase font-bold tracking-wider mb-1">Total Other Cost (BDT)</p>
                <p className="text-2xl font-bold text-slate-700">{fmt(calculations.totalOtherCost)}</p>
              </div>
              <div className="bg-slate-900 p-4 rounded-xl shadow-sm border border-slate-900">
                <p className="text-xs text-slate-400 uppercase font-bold tracking-wider mb-1">Grand Total (BDT)</p>
                <p className="text-2xl font-bold text-white">{fmt(calculations.totalGrossCost)}</p>
              </div>
            </div>

            {/* Table Card */}
            <div className="bg-white rounded-2xl shadow-sm border border-slate-200 flex-1 overflow-hidden flex flex-col">
              <div className="px-6 py-4 border-b border-slate-200 flex justify-between items-center bg-slate-50">
                <h2 className="font-semibold text-slate-700 flex items-center gap-2">
                  <Calculator className="w-4 h-4" />
                  Cost Invoice
                </h2>
                <div className="flex items-center gap-3">
                  <span className="text-xs font-medium px-2 py-1 bg-blue-100 text-blue-700 rounded">
                    {products.length} Items
                  </span>

                  {/* Export Dropdown */}
                  <div className="relative">
                    <button
                      onClick={() => setShowExportMenu(!showExportMenu)}
                      className="flex items-center gap-1 text-xs font-medium text-slate-600 hover:text-blue-600 bg-slate-100 hover:bg-blue-50 px-2 py-1 rounded transition-colors"
                    >
                      <Download className="w-3 h-3" /> Export
                    </button>

                    {showExportMenu && (
                      <>
                        <div className="fixed inset-0 z-10" onClick={() => setShowExportMenu(false)}></div>
                        <div className="absolute right-0 top-full mt-1 w-32 bg-white rounded-lg shadow-lg border border-slate-100 py-1 z-20 flex flex-col cursor-pointer">
                          <button onClick={exportToExcel} className="text-left px-3 py-2 text-xs text-slate-700 hover:bg-slate-50 hover:text-blue-600 transition-colors">
                            Download Excel
                          </button>
                          <button onClick={exportToCSV} className="text-left px-3 py-2 text-xs text-slate-700 hover:bg-slate-50 hover:text-blue-600 transition-colors">
                            Download CSV
                          </button>
                        </div>
                      </>
                    )}
                  </div>
                </div>
              </div>

              <div className="overflow-x-auto flex-1">
                {products.length === 0 ? (
                  <div className="flex flex-col items-center justify-center h-64 text-slate-400">
                    <Package className="w-12 h-12 mb-3 opacity-20" />
                    <p>No products added yet.</p>
                    <p className="text-sm">Add items using the form to see calculations.</p>
                  </div>
                ) : (
                  <table className="w-full text-left border-collapse">
                    <thead>
                      <tr className="bg-slate-50 text-[10px] md:text-xs text-slate-500 uppercase tracking-wider border-b border-slate-200">
                        <th className="px-2 py-3 font-semibold whitespace-nowrap">Product</th>
                        <th className="px-2 py-3 font-semibold text-center whitespace-nowrap">Qty</th>
                        <th className="px-2 py-3 font-semibold text-right whitespace-nowrap">Unit Price<br /><span className="normal-case text-[10px] text-slate-400">(RMB)</span></th>
                        <th className="px-2 py-3 font-semibold text-right whitespace-nowrap">Total Price<br /><span className="normal-case text-[10px] text-slate-400">(RMB)</span></th>
                        <th className="px-2 py-3 font-semibold text-right whitespace-nowrap">Gross Weight<br /><span className="normal-case text-[10px] text-slate-400">(Net + Box Wgt)</span></th>
                        <th className="px-2 py-3 font-semibold text-right whitespace-nowrap">Shipment Cost<br /><span className="normal-case text-[10px] text-slate-400">(Unit)</span></th>
                        <th className="px-2 py-3 font-semibold text-right whitespace-nowrap">Extra Cost<br /><span className="normal-case text-[10px] text-slate-400">(Others)</span></th>
                        <th className="px-2 py-3 font-semibold text-right whitespace-nowrap">Total Cost</th>
                        <th className="px-2 py-3 font-semibold text-right whitespace-nowrap">Per Unit Price</th>
                        <th className="px-1 py-3"></th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100 text-xs md:text-sm">
                      {calculations.processedProducts.map((p) => (
                        <tr key={p.id} className="hover:bg-slate-50 group transition-colors">
                          <td className="px-2 py-2">
                            <div className="flex flex-col">
                              <div className="flex items-center gap-2">
                                <span className="text-xs text-slate-400 font-mono bg-slate-100 px-1 rounded">{p.sn}</span>
                                <span className="font-medium text-slate-800">{p.name}</span>
                              </div>
                            </div>
                          </td>
                          <td className="px-2 py-2 text-center text-slate-600">
                            {p.quantity}
                          </td>
                          <td className="px-2 py-2 text-right">
                            <span className="text-slate-600 font-medium">{p.cost}</span>
                          </td>
                          <td className="px-2 py-2 text-right">
                            <span className="text-emerald-600 font-medium">{fmt(p.totalRmbParams)}</span>
                          </td>
                          <td className="px-2 py-2 text-right">
                            <div className="flex flex-col items-end">
                              <span className="font-medium text-slate-700">{fmtW(p.unitGrossWeight)} kg</span>
                            </div>
                          </td>
                          <td className="px-2 py-2 text-right">
                            <span className="text-slate-700 font-medium">{fmt(p.totalShippingCost)}</span>
                          </td>
                          <td className="px-2 py-2 text-right">
                            <span className="text-slate-700 font-medium">{fmt(p.allocatedOtherCost)}</span>
                          </td>
                          <td className="px-2 py-2 text-right font-bold text-slate-800">
                            {fmt(p.totalLineCost)}
                          </td>
                          <td className="px-2 py-2 text-right bg-blue-50/50 font-bold text-blue-700">
                            {fmt(p.unitGrossCost)}
                          </td>
                          <td className="px-1 py-2 text-right">
                            <div className="flex items-center justify-end gap-1">
                              <button
                                onClick={() => startEdit(p)}
                                className="text-slate-300 hover:text-blue-500 p-1 rounded transition-colors"
                                title="Edit"
                              >
                                <Edit2 className="w-3 h-3" />
                              </button>
                              <button
                                onClick={() => removeProduct(p.id)}
                                className="text-slate-300 hover:text-red-500 p-1 rounded transition-colors"
                                title="Remove"
                              >
                                <Trash2 className="w-3 h-3" />
                              </button>
                            </div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                    <tfoot className="bg-slate-50 border-t border-slate-200">
                      <tr>
                        <td className="px-2 py-3 font-bold text-slate-600">Totals</td>
                        <td colSpan="2"></td>
                        <td className="px-2 py-3 text-right font-bold text-emerald-700">{fmt(calculations.totalRmbParams)}</td>
                        <td className="px-2 py-3 text-right font-bold text-slate-600">{fmtW(globalSettings.totalBoxWeight)} kg</td>
                        <td className="px-2 py-3 text-right font-bold text-slate-600">{fmt(calculations.totalShippingCost)}</td>
                        <td colSpan="1"></td>
                        <td className="px-2 py-3 text-right font-bold text-xl text-slate-900">{fmt(calculations.totalGrossCost)}</td>
                        <td colSpan="2"></td>
                      </tr>
                    </tfoot>
                  </table>
                )}
              </div>
            </div>

            {/* Explainer / Legend */}
            <div className="bg-blue-50 border border-blue-100 rounded-xl p-4 flex gap-4 text-sm text-blue-800">
              <Info className="w-5 h-5 flex-shrink-0 mt-0.5" />
              <div className="space-y-1">
                <p className="font-semibold">Shipping Logic:</p>
                <ul className="list-disc pl-4 space-y-1 opacity-80">
                  <li>Supplier to CN Wirehouse is now a <strong>Fixed Box Cost</strong> (RMB), distributed by weight.</li>
                  <li>CN to BD costs are calculated directly in BDT per KG.</li>
                  <li>All shipping costs are based on the Unit Gross Weight (Product Weight + Proportionate Box Weight).</li>
                </ul>
              </div>
            </div>

          </div>
        </div>
      </div>
    </div>
  );
};

export default App;
