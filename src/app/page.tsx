"use client";

import { useState, FormEvent, ChangeEvent } from "react";
import Image from "next/image";

export default function Home() {
  const [loading, setLoading] = useState(false);
  const [file, setFile] = useState<File | null>(null);
  const [logoTop, setLogoTop] = useState<string | null>("/logo.svg");
  const [logoBottom, setLogoBottom] = useState<string | null>("/logo.svg");

  // Configuration State
  const [config, setConfig] = useState({
    prefix: "MC",
    startInv: 1,
    date: new Date().toISOString().split("T")[0],
    valid: new Date(new Date().setMonth(new Date().getMonth() + 3))
      .toISOString()
      .split("T")[0],
    title: "NEW YEAR COMPLEMENTARY",
    greeting: "HAPPY NEW YEAR 2025!",
    terms: "This invoice can only be claimed at the Factory.",
    company: "Perera and Sons Bakers (Pvt) Ltd",
    address: "122-124, M D H Jayawardena Mawatha, Rajagiriya",
    item1: "VANILLA CAKE 1KG",
    price1: 1500,
    qty1: 1,
    item2: "VANILLA CAKE 400G",
    price2: 700,
    qty2: 1,
  });

  const handleChange = (
    e: ChangeEvent<HTMLInputElement | HTMLTextAreaElement>
  ) => {
    const { name, value } = e.target;
    setConfig((prev) => ({ ...prev, [name]: value }));
  };

  const handleFileChange = (e: ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      setFile(e.target.files[0]);
    }
  };

  const handleLogoTopChange = (e: ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      setLogoTop(URL.createObjectURL(e.target.files[0]));
    }
  };

  const handleLogoBottomChange = (e: ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      setLogoBottom(URL.createObjectURL(e.target.files[0]));
    }
  };

  const handleSubmit = async (e: FormEvent) => {
    e.preventDefault();
    if (!file) {
      alert("Please upload an Excel file first!");
      return;
    }

    setLoading(true);
    const formData = new FormData();
    formData.append("file", file);
    formData.append("config", JSON.stringify(config));

    // We need to get the actual file objects for logos if they were changed
    const logoTopInput = document.querySelector(
      'input[name="logoTop"]'
    ) as HTMLInputElement;
    if (logoTopInput?.files?.[0]) {
      formData.append("logoTop", logoTopInput.files[0]);
    }

    const logoBottomInput = document.querySelector(
      'input[name="logoBottom"]'
    ) as HTMLInputElement;
    if (logoBottomInput?.files?.[0]) {
      formData.append("logoBottom", logoBottomInput.files[0]);
    }

    try {
      const response = await fetch("/api/generate", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        const err = await response.json();
        throw new Error(err.message || "Failed to generate");
      }

      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "Professional_Invoices.xlsx";
      document.body.appendChild(a);
      a.click();
      a.remove();
    } catch (error: any) {
      alert("Error: " + error.message);
    } finally {
      setLoading(false);
    }
  };

  const total1 = Number(config.price1) * Number(config.qty1);
  const total2 = Number(config.price2) * Number(config.qty2);
  const grandTotal = total1 + total2;

  return (
    <main className="min-h-screen bg-gray-100 flex flex-col lg:flex-row">
      {/* Sidebar Controls */}
      <div className="w-full lg:w-80 bg-white border-r border-gray-200 p-6 flex flex-col gap-6 shadow-lg z-10 h-auto lg:h-screen overflow-y-auto">
        <div>
          <h2 className="text-xl font-bold text-slate-800 mb-4 flex items-center gap-2">
            <span className="bg-black text-white p-1 rounded">P&S</span>{" "}
            Generator
          </h2>
          <p className="text-sm text-gray-500">
            Configure your invoice template and upload employee data.
          </p>
        </div>

        <div className="space-y-4">
          <div className="p-4 border-2 border-dashed border-gray-300 rounded-xl bg-gray-50">
            <label className="block text-sm font-bold text-gray-900 mb-2">
              1. Upload Data (Excel)
            </label>
            <input
              type="file"
              accept=".xlsx, .xls"
              onChange={handleFileChange}
              className="block w-full text-sm text-slate-500 file:mr-4 file:py-2 file:px-4 file:rounded-full file:border-0 file:text-xs file:font-semibold file:bg-gray-800 file:text-white hover:file:bg-gray-700 cursor-pointer"
            />
            <p className="text-xs text-gray-600 mt-2">
              Columns: NAME, EPF, DEPARTMENT
            </p>
          </div>

          <div className="space-y-3">
            <label className="block text-sm font-semibold text-gray-700">
              Logos
            </label>
            <div className="flex gap-2">
              <div className="flex-1">
                <label className="text-xs text-gray-500 block mb-1">
                  Top Logo
                </label>
                <input
                  type="file"
                  name="logoTop"
                  accept="image/*"
                  onChange={handleLogoTopChange}
                  className="text-xs w-full"
                />
              </div>
            </div>
            <div className="flex gap-2">
              <div className="flex-1">
                <label className="text-xs text-gray-500 block mb-1">
                  Bottom Logo
                </label>
                <input
                  type="file"
                  name="logoBottom"
                  accept="image/*"
                  onChange={handleLogoBottomChange}
                  className="text-xs w-full"
                />
              </div>
            </div>
          </div>

          <button
            onClick={handleSubmit}
            disabled={loading}
            className="w-full py-3 px-4 bg-black hover:bg-gray-800 text-white font-bold rounded-lg shadow-md transition-all transform hover:scale-[1.02] disabled:opacity-50 disabled:cursor-not-allowed flex justify-center items-center gap-2"
          >
            {loading ? (
              <>
                <div className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin"></div>
                Processing...
              </>
            ) : (
              <>
                <span>Generate Invoices</span>
                <svg
                  className="w-4 h-4"
                  fill="none"
                  stroke="currentColor"
                  viewBox="0 0 24 24"
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    strokeWidth="2"
                    d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"
                  ></path>
                </svg>
              </>
            )}
          </button>
        </div>

        <div className="mt-auto pt-6 border-t border-gray-100">
          <p className="text-xs text-center text-gray-400">
            Developed by{" "}
            <span className="font-bold text-gray-600">Flowiix (pvt) LTD</span>
          </p>
        </div>
      </div>

      {/* Preview Area */}
      <div className="flex-1 p-8 overflow-auto flex justify-center bg-gray-100">
        <div className="bg-white shadow-2xl w-[210mm] min-h-[297mm] flex relative">
          {/* Left Strip */}
          <div className="w-10 bg-black flex flex-col items-center justify-center text-white shrink-0">
            <div
              className="transform -rotate-90 whitespace-nowrap font-bold tracking-widest text-xs opacity-80"
              style={{ writingMode: "vertical-rl" }}
            >
              {config.company.toUpperCase()} - OFFICIAL COPY
            </div>
          </div>

          {/* Main Content */}
          <div className="flex-1 p-8 flex flex-col relative">
            {/* Header */}
            <div className="text-center mb-6">
              <input
                type="text"
                name="company"
                value={config.company}
                onChange={handleChange}
                className="w-full text-center text-xl font-bold text-black uppercase border-none focus:ring-0 p-0 bg-transparent placeholder-gray-400"
                placeholder="COMPANY NAME"
              />
              <input
                type="text"
                name="address"
                value={config.address}
                onChange={handleChange}
                className="w-full text-center text-sm text-gray-600 border-none focus:ring-0 p-0 bg-transparent mt-1"
                placeholder="Company Address"
              />
            </div>

            {/* Title */}
            <div className="mb-4 text-center">
              <input
                type="text"
                name="title"
                value={config.title}
                onChange={handleChange}
                className="w-full text-center font-bold text-gray-800 border-b-2 border-black pb-1 focus:ring-0 bg-transparent"
              />
            </div>

            {/* Info Bar */}
            <div className="bg-gray-100 p-2 mb-6 flex justify-center items-center gap-4 border border-gray-200">
              <div className="flex items-center gap-2">
                <span className="text-sm font-bold text-gray-700">
                  INVOICE NO:
                </span>
                <div className="flex items-center">
                  <input
                    type="text"
                    name="prefix"
                    value={config.prefix}
                    onChange={handleChange}
                    className="w-12 text-sm font-mono border-gray-300 rounded px-1 py-0.5 h-6"
                  />
                  <input
                    type="number"
                    name="startInv"
                    value={config.startInv}
                    onChange={handleChange}
                    className="w-16 text-sm font-mono border-gray-300 rounded px-1 py-0.5 h-6 ml-1"
                  />
                </div>
              </div>
              <span className="text-gray-400">|</span>
              <div className="flex items-center gap-2">
                <span className="text-sm font-bold text-gray-700">DATE:</span>
                <input
                  type="date"
                  name="date"
                  value={config.date}
                  onChange={handleChange}
                  className="text-sm border-gray-300 rounded px-1 py-0.5 h-6"
                />
              </div>
            </div>

            {/* Employee Placeholder */}
            <div className="mb-8 space-y-2 pl-4 border-l-4 border-black bg-gray-50 p-4 rounded-r">
              <div className="grid grid-cols-[120px_1fr] gap-2 items-center">
                <span className="text-sm font-bold text-black">NAME</span>
                <span className="text-sm text-gray-600 font-mono">
                  : [Employee Name from Excel]
                </span>
              </div>
              <div className="grid grid-cols-[120px_1fr] gap-2 items-center">
                <span className="text-sm font-bold text-black">EPF NO</span>
                <span className="text-sm text-gray-600 font-mono">
                  : [EPF from Excel]
                </span>
              </div>
              <div className="grid grid-cols-[120px_1fr] gap-2 items-center">
                <span className="text-sm font-bold text-black">DEPARTMENT</span>
                <span className="text-sm text-gray-600 font-mono">
                  : [Dept from Excel]
                </span>
              </div>
            </div>

            {/* Table */}
            <div className="mb-8">
              <div className="grid grid-cols-12 bg-black text-white text-sm font-bold py-2 px-2 rounded-t">
                <div className="col-span-6">DESCRIPTION</div>
                <div className="col-span-2 text-center">QTY</div>
                <div className="col-span-2 text-right">PRICE</div>
                <div className="col-span-2 text-right">TOTAL</div>
              </div>

              {/* Row 1 */}
              <div className="grid grid-cols-12 border-b border-gray-200 py-2 px-2 items-center hover:bg-gray-50">
                <div className="col-span-6">
                  <input
                    type="text"
                    name="item1"
                    value={config.item1}
                    onChange={handleChange}
                    className="w-full text-sm border-none bg-transparent focus:ring-0 p-0 font-medium"
                    placeholder="Item Description"
                  />
                </div>
                <div className="col-span-2">
                  <input
                    type="number"
                    name="qty1"
                    value={config.qty1}
                    onChange={handleChange}
                    className="w-full text-center text-sm border-gray-200 rounded p-1"
                  />
                </div>
                <div className="col-span-2">
                  <input
                    type="number"
                    name="price1"
                    value={config.price1}
                    onChange={handleChange}
                    className="w-full text-right text-sm border-gray-200 rounded p-1"
                  />
                </div>
                <div className="col-span-2 text-right text-sm font-mono">
                  {total1.toLocaleString("en-US", {
                    style: "currency",
                    currency: "LKR",
                  })}
                </div>
              </div>

              {/* Row 2 */}
              <div className="grid grid-cols-12 border-b border-gray-200 py-2 px-2 items-center hover:bg-gray-50">
                <div className="col-span-6">
                  <input
                    type="text"
                    name="item2"
                    value={config.item2}
                    onChange={handleChange}
                    className="w-full text-sm border-none bg-transparent focus:ring-0 p-0 font-medium"
                    placeholder="Item Description"
                  />
                </div>
                <div className="col-span-2">
                  <input
                    type="number"
                    name="qty2"
                    value={config.qty2}
                    onChange={handleChange}
                    className="w-full text-center text-sm border-gray-200 rounded p-1"
                  />
                </div>
                <div className="col-span-2">
                  <input
                    type="number"
                    name="price2"
                    value={config.price2}
                    onChange={handleChange}
                    className="w-full text-right text-sm border-gray-200 rounded p-1"
                  />
                </div>
                <div className="col-span-2 text-right text-sm font-mono">
                  {total2.toLocaleString("en-US", {
                    style: "currency",
                    currency: "LKR",
                  })}
                </div>
              </div>

              {/* Grand Total */}
              <div className="flex justify-end mt-4">
                <div className="bg-black text-white p-4 rounded-lg shadow-md flex gap-4 items-center">
                  <span className="text-sm font-bold uppercase">
                    Grand Total
                  </span>
                  <span className="text-2xl font-bold">
                    {grandTotal.toLocaleString("en-US", {
                      style: "currency",
                      currency: "LKR",
                    })}
                  </span>
                </div>
              </div>
            </div>

            {/* Footer Info */}
            <div className="mt-auto">
              <div className="flex justify-center mb-8">
                <div className="bg-gray-100 px-6 py-2 rounded-full border border-gray-200 flex items-center gap-2">
                  <span className="text-sm font-bold text-gray-600">
                    VALID UNTIL:
                  </span>
                  <input
                    type="date"
                    name="valid"
                    value={config.valid}
                    onChange={handleChange}
                    className="bg-transparent border-none text-sm font-bold text-gray-800 focus:ring-0 p-0"
                  />
                </div>
              </div>

              <div className="flex justify-between items-end mb-8">
                <div className="w-32 h-16 relative">
                  {logoBottom && (
                    <Image
                      src={logoBottom}
                      alt="Logo"
                      fill
                      className="object-contain object-left-bottom"
                    />
                  )}
                </div>
                <div className="text-right">
                  <div className="w-48 border-b border-gray-400 mb-2"></div>
                  <p className="text-xs font-bold text-gray-600 uppercase">
                    Authorized Signature
                  </p>
                </div>
              </div>

              <div className="text-center border-t border-gray-200 pt-4">
                <textarea
                  name="terms"
                  value={config.terms}
                  onChange={handleChange}
                  rows={2}
                  className="w-full text-center text-xs text-gray-500 italic bg-transparent border-none focus:ring-0 resize-none"
                />
              </div>
            </div>

            {/* Floating Logo Top */}
            {logoTop && (
              <div className="absolute top-8 left-8 w-20 h-20">
                <Image
                  src={logoTop}
                  alt="Logo"
                  fill
                  className="object-contain"
                />
              </div>
            )}
          </div>
        </div>
      </div>
    </main>
  );
}
