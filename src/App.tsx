import React, { useState, useMemo, useRef, useEffect } from "react";
import {
  Upload,
  Download,
  Search,
  Plus,
  Trash2,
  Edit2,
  Check,
  CheckSquare,
  Square,
  Printer,
  Settings,
  Building,
  MapPin,
  RotateCcw,
  FileSpreadsheet,
  Sliders,
  X,
  Eye,
  HelpCircle,
  Clock,
  Phone,
  Grid,
  Info,
  Sparkles
} from "lucide-react";
import * as XLSX from "xlsx";
import { School, SenderInfo, PrinterPreset } from "./types";
import { defaultSchools } from "./defaultSchools";

const BUILTIN_PRESETS: PrinterPreset[] = [
  {
    id: "preset_hp_laserjet",
    name: "HP LaserJet Standard (DL Envelope)",
    toBlockTopPadding: 15,
    toBlockXShift: 0,
    toBlockWidth: 138,
    isBuiltIn: true
  },
  {
    id: "preset_epson_inkjet",
    name: "Epson Inkjet (DL - Left Shift Offset)",
    toBlockTopPadding: 20,
    toBlockXShift: 5,
    toBlockWidth: 130,
    isBuiltIn: true
  },
  {
    id: "preset_canon_pixma",
    name: "Canon PIXMA (Compact DL - Centered)",
    toBlockTopPadding: 12,
    toBlockXShift: -10,
    toBlockWidth: 145,
    isBuiltIn: true
  },
  {
    id: "preset_windowed_standard",
    name: "Windowed DL Style (Low Centered)",
    toBlockTopPadding: 25,
    toBlockXShift: 12,
    toBlockWidth: 120,
    isBuiltIn: true
  }
];

export default function App() {
  // State for original load of default school data
  const [schools, setSchools] = useState<School[]>(() => {
    const saved = localStorage.getItem("b2p_schools_database");
    if (saved) {
      try {
        return JSON.parse(saved);
      } catch (e) {
        return defaultSchools;
      }
    }
    return defaultSchools;
  });

  // State for multiple sender company profiles
  const [senders, setSenders] = useState<SenderInfo[]>(() => {
    const saved = localStorage.getItem("b2p_senders_list");
    if (saved) {
      try {
        return JSON.parse(saved);
      } catch (e) {
        // fallback
      }
    }
    return [
      {
        id: "1",
        name: "B2P INTERNATIONAL",
        address: "10/15 DEVASSY SQUARE, MARATHAKAM CENTRE,\nPURANATTUKARA, THRISSUR, KERALA - 680551\nPH: +91 85899 09034"
      }
    ];
  });

  const [activeSenderId, setActiveSenderId] = useState<string>(() => {
    const savedId = localStorage.getItem("b2p_active_sender_id");
    return savedId || "1";
  });

  // Derived active sender profile
  const activeSender = useMemo(() => {
    return senders.find((s) => s.id === activeSenderId) || senders[0] || {
      id: "1",
      name: "B2P INTERNATIONAL",
      address: "10/15 DEVASSY SQUARE, MARATHAKAM CENTRE,\nPURANATTUKARA, THRISSUR, KERALA - 680551\nPH: +91 85899 09034"
    };
  }, [senders, activeSenderId]);

  // UI state managers
  const [activeTab, setActiveTab] = useState<"database" | "preview" | "settings">("database");
  const [searchQuery, setSearchQuery] = useState("");
  const [selectedSchoolId, setSelectedSchoolId] = useState<string | null>(null);

  // Custom high-fidelity modal dialog system to bypass iframe sandbox browser blocks (alert/confirm/prompt)
  interface DialogConfig {
    isOpen: boolean;
    type: "confirm" | "prompt" | "double_prompt" | "alert";
    title: string;
    message: string;
    inputValue?: string;
    inputValueSecond?: string;
    inputLabel?: string;
    inputLabelSecond?: string;
    onConfirm: (val1?: string, val2?: string) => void;
  }

  const [dialog, setDialog] = useState<DialogConfig>({
    isOpen: false,
    type: "confirm",
    title: "",
    message: "",
    onConfirm: () => {}
  });

  const [dialogInput1, setDialogInput1] = useState<string>("");
  const [dialogInput2, setDialogInput2] = useState<string>("");

  const triggerConfirm = (title: string, message: string, onConfirm: () => void) => {
    setDialog({
      isOpen: true,
      type: "confirm",
      title,
      message,
      onConfirm
    });
  };

  const triggerPrompt = (
    title: string,
    message: string,
    defaultValue: string,
    inputLabel: string,
    onConfirm: (val: string) => void
  ) => {
    setDialogInput1(defaultValue);
    setDialog({
      isOpen: true,
      type: "prompt",
      title,
      message,
      inputLabel,
      onConfirm: (val1) => onConfirm(val1 || "")
    });
  };

  const triggerDoublePrompt = (
    title: string,
    message: string,
    val1: string,
    lbl1: string,
    val2: string,
    lbl2: string,
    onConfirm: (res1: string, res2: string) => void
  ) => {
    setDialogInput1(val1);
    setDialogInput2(val2);
    setDialog({
      isOpen: true,
      type: "double_prompt",
      title,
      message,
      inputLabel: lbl1,
      inputLabelSecond: lbl2,
      onConfirm: (res1, res2) => onConfirm(res1 || "", res2 || "")
    });
  };

  const triggerAlert = (title: string, message: string) => {
    setDialog({
      isOpen: true,
      type: "alert",
      title,
      message,
      onConfirm: () => {}
    });
  };

  // Styling preference state
  const [letterSpacing, setLetterSpacing] = useState<number>(5); // 0px to 10px spacing
  const [fromLetterSpacing, setFromLetterSpacing] = useState<number>(3);
  const [showToPrefix, setShowToPrefix] = useState<boolean>(true);
  const [stampText, setStampText] = useState<string>("AFFIX\nPOSTAGE\nSTAMP\nHERE");
  const [enableStampBorder, setEnableStampBorder] = useState<boolean>(true);

  // Controls for interactive workspace viewing
  const [showCalibrationSidebar, setShowCalibrationSidebar] = useState<boolean>(true);
  const [previewZoom, setPreviewZoom] = useState<number>(75); // zoom percentage from 30% to 100%

  // New states for unified font and TO area alignment calibration
  const [envelopeFont, setEnvelopeFont] = useState<string>(() => {
    return localStorage.getItem("b2p_envelope_font") || "Arial, Helvetica, sans-serif";
  });
  const [toBlockTopPadding, setToBlockTopPadding] = useState<number>(() => {
    const saved = localStorage.getItem("b2p_to_top_padding");
    return saved ? Number(saved) : 15;
  });
  const [toBlockXShift, setToBlockXShift] = useState<number>(() => {
    const saved = localStorage.getItem("b2p_to_x_shift");
    return saved ? Number(saved) : 0;
  });
  const [toBlockWidth, setToBlockWidth] = useState<number>(() => {
    const saved = localStorage.getItem("b2p_to_block_width");
    return saved ? Number(saved) : 138;
  });

  // State for printer / alignment presets
  const [presets, setPresets] = useState<PrinterPreset[]>(() => {
    const saved = localStorage.getItem("b2p_printer_presets");
    if (saved) {
      try {
        const parsed = JSON.parse(saved) as PrinterPreset[];
        const customs = parsed.filter(p => !p.isBuiltIn);
        return [...BUILTIN_PRESETS, ...customs];
      } catch (e) {
        return BUILTIN_PRESETS;
      }
    }
    return BUILTIN_PRESETS;
  });

  const [selectedPresetId, setSelectedPresetId] = useState<string>(() => {
    return localStorage.getItem("b2p_selected_preset_id") || "preset_hp_laserjet";
  });

  // Bidirectional matching effect: checks if current sliders match any preset
  useEffect(() => {
    const matchingPreset = presets.find(
      (p) =>
        p.toBlockTopPadding === toBlockTopPadding &&
        p.toBlockXShift === toBlockXShift &&
        p.toBlockWidth === toBlockWidth
    );
    if (matchingPreset) {
      if (selectedPresetId !== matchingPreset.id) {
        setSelectedPresetId(matchingPreset.id);
      }
    } else {
      if (selectedPresetId !== "custom") {
        setSelectedPresetId("custom");
      }
    }
  }, [toBlockTopPadding, toBlockXShift, toBlockWidth, presets, selectedPresetId]);

  const isInsideIframe = useMemo(() => {
    try {
      return window.self !== window.top;
    } catch (e) {
      return true;
    }
  }, []);
  
  // School Modal editing / manual entry state
  const [editingSchool, setEditingSchool] = useState<Partial<School> | null>(null);
  const [isAddingNew, setIsAddingNew] = useState(false);
  const [newSchoolName, setNewSchoolName] = useState("");
  const [newSchoolAddr, setNewSchoolAddr] = useState("");

  // File Upload mapping helper
  const [excelMapping, setExcelMapping] = useState<{
    show: boolean;
    columns: string[];
    data: any[];
    nameKey: string;
    addrKey: string;
  } | null>(null);

  // Search input and list filters
  const fileInputRef = useRef<HTMLInputElement>(null);
  const logoInputRef = useRef<HTMLInputElement>(null);

  // Sync state to local storage on changes to avoid data loss
  useEffect(() => {
    localStorage.setItem("b2p_schools_database", JSON.stringify(schools));
  }, [schools]);

  useEffect(() => {
    localStorage.setItem("b2p_senders_list", JSON.stringify(senders));
  }, [senders]);

  useEffect(() => {
    localStorage.setItem("b2p_active_sender_id", activeSenderId);
  }, [activeSenderId]);

  useEffect(() => {
    localStorage.setItem("b2p_envelope_font", envelopeFont);
  }, [envelopeFont]);

  useEffect(() => {
    localStorage.setItem("b2p_to_top_padding", String(toBlockTopPadding));
  }, [toBlockTopPadding]);

  useEffect(() => {
    localStorage.setItem("b2p_to_x_shift", String(toBlockXShift));
  }, [toBlockXShift]);

  useEffect(() => {
    localStorage.setItem("b2p_to_block_width", String(toBlockWidth));
  }, [toBlockWidth]);

  useEffect(() => {
    localStorage.setItem("b2p_printer_presets", JSON.stringify(presets));
  }, [presets]);

  useEffect(() => {
    localStorage.setItem("b2p_selected_preset_id", selectedPresetId);
  }, [selectedPresetId]);

  // Filtered Schools calculation
  const filteredSchools = useMemo(() => {
    const query = searchQuery.trim().toLowerCase();
    if (!query) return schools;
    return schools.filter(
      (s) =>
        s.name.toLowerCase().includes(query) ||
        s.address.toLowerCase().includes(query)
    );
  }, [schools, searchQuery]);

  // Selected for printing
  const selectedCount = useMemo(() => {
    return schools.filter((s) => s.selected).length;
  }, [schools]);

  // Bulk actions
  const handleToggleSelectAll = () => {
    const allSelected = filteredSchools.every((s) => s.selected);
    const updatedIds = new Set(filteredSchools.map((s) => s.id));
    setSchools((prev) =>
      prev.map((s) => {
        if (updatedIds.has(s.id)) {
          return { ...s, selected: !allSelected };
        }
        return s;
      })
    );
  };

  const handleToggleSingleSelect = (id: string) => {
    setSchools((prev) =>
      prev.map((s) => {
        if (s.id === id) {
          return { ...s, selected: !s.selected };
        }
        return s;
      })
    );
  };

  // Add / Edit actions
  const handleStartEdit = (school: School) => {
    setEditingSchool({ ...school });
  };

  const handleSaveEdit = () => {
    if (!editingSchool || !editingSchool.id) return;
    setSchools((prev) =>
      prev.map((s) =>
        s.id === editingSchool.id
          ? {
              ...s,
              name: editingSchool.name || "",
              address: editingSchool.address || ""
            }
          : s
      )
    );
    setEditingSchool(null);
  };

  const autoFormatAddressText = (text: string): string => {
    if (!text) return "";
    
    // Split into lines to retain structure
    const lines = text.split("\n");
    
    const formattedLines = lines.map((line) => {
      const words = line.trim().split(/\s+/);
      if (words.length === 1 && words[0] === "") return "";

      const upperCaseAbbrs = ["po", "p.o.", "pin", "ph", "mla", "cms", "co", "io", "sl", "dl", "hp", "uk"];
      const lowercaseConjs = ["and", "or", "of", "the", "in", "at", "by", "via", "for", "with", "to", "on"];
      
      const formattedWords = words.map((word, wordIndex) => {
        // Keep pure numeric coordinates or phone country codes/numbers intact
        if (/^\+?\d+[\/\d\-\+]*$/.test(word)) {
          return word;
        }
        if (!/[a-zA-Z]/.test(word)) {
          return word;
        }
        
        const cleaned = word.toLowerCase().replace(/[^a-z]/g, "");
        if (upperCaseAbbrs.includes(cleaned)) {
          return word.toUpperCase();
        }
        
        let newWord = word.toLowerCase();
        if (wordIndex > 0 && lowercaseConjs.includes(cleaned)) {
          return newWord;
        }
        
        return newWord.replace(/(^[a-z]|(?<=\W)[a-z])/g, (c) => c.toUpperCase());
      });
      
      let formattedLine = formattedWords.join(" ");
      
      // Ensure zip code spacing consistency using simple regex.
      // Standardize 6-digit pin codes with consistent spacing (e.g., 680551 -> 680 551)
      formattedLine = formattedLine.replace(/\b(\d{3})\s*(\d{3})\b/g, "$1 $2");
      // Standardize US 5-4 digits ZIP code dashes
      formattedLine = formattedLine.replace(/\b(\d{5})\s*-\s*(\d{4})\b/g, "$1-$2");
      
      return formattedLine;
    });
    
    return formattedLines.join("\n");
  };

  const handleAddNewSchool = () => {
    if (!newSchoolName.trim() || !newSchoolAddr.trim()) return;
    const newSchool: School = {
      id: "school_" + Date.now().toString(),
      name: newSchoolName.trim(),
      address: newSchoolAddr.trim(),
      selected: true
    };
    setSchools((prev) => [newSchool, ...prev]);
    setNewSchoolName("");
    setNewSchoolAddr("");
    setIsAddingNew(false);
    setActiveTab("database");
  };

  const handleDeleteSchool = (id: string) => {
    const school = schools.find((s) => s.id === id);
    const displayName = school ? school.name : "this address";
    triggerConfirm(
      "Delete Address Recipient",
      `Are you sure you want to delete the address record for "${displayName}"? This will remove the entry from your database.`,
      () => {
        setSchools((prev) => prev.filter((s) => s.id !== id));
        if (selectedSchoolId === id) setSelectedSchoolId(null);
      }
    );
  };

  const handleResetToPresets = () => {
    triggerConfirm(
      "Restore Database Defaults",
      "Are you sure you want to restore the database to the 40 original preview schools? Any newly added, edited, or uploaded school records in this session will be permanently deleted.",
      () => {
        setSchools(defaultSchools);
        localStorage.removeItem("b2p_schools_database");
      }
    );
  };

  const handleDownloadTemplate = () => {
    try {
      const templateData = [
        {
          "School Name": "Sree Rama Convent Senior Secondary School",
          "Recipient Address": "Ashram Lane, Punkunnam\nThrissur, Kerala Pin - 680002"
        },
        {
          "School Name": "Amala Mary Rani Public School",
          "Recipient Address": "Puranattukara P.O., Thrissur\nKerala - 680551\nIndia"
        },
        {
          "School Name": "St. Paul's Public School",
          "Recipient Address": "St. Paul's Street, Kuriachira\nThrissur, Kerala Pin - 680006"
        }
      ];
      
      const worksheet = XLSX.utils.json_to_sheet(templateData);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "Recipient Roster");
      
      const excelBuffer = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
      const dataBlob = new Blob([excelBuffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      
      const url = window.URL.createObjectURL(dataBlob);
      const link = document.createElement("a");
      link.href = url;
      link.setAttribute("download", "Envelope_Addresses_Template.xlsx");
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    } catch (err) {
      alert("Could not generate Excel template. Please try refreshing.");
    }
  };

  // Excel / CSV File Import Parser using 'xlsx' package
  const handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: "array" });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Parse raw rows as arrays or objects
        const rawRows = XLSX.utils.sheet_to_json<any>(worksheet, { header: 1 });
        if (rawRows.length === 0) {
          alert("Your uploaded spreadsheet is empty!");
          return;
        }

        // Detect column header candidates from the first non-empty row
        let headerRowIndex = 0;
        for (let idx = 0; idx < rawRows.length; idx++) {
          if (Array.isArray(rawRows[idx]) && rawRows[idx].length > 0) {
            headerRowIndex = idx;
            break;
          }
        }

        const rawColumns: string[] = rawRows[headerRowIndex].map((col: any) =>
          String(col || "").trim()
        );

        // Convert the rest to JSON objects with detected column mappings
        const listData = XLSX.utils.sheet_to_json<any>(worksheet, { defval: "" });

        // Auto-guess columns
        let guessNameKey = "";
        let guessAddrKey = "";

        // Common variants of header names
        const nameVariants = ["school", "school name", "name", "recipient", "to", "organisation", "organization", "title", "customer"];
        const addrVariants = ["address", "addr", "recipient address", "school address", "street", "city", "location", "address details", "postal"];

        rawColumns.forEach((col) => {
          const lower = col.toLowerCase();
          if (nameVariants.some((v) => lower.includes(v)) && !guessNameKey) {
            guessNameKey = col;
          }
          if (addrVariants.some((v) => lower.includes(v)) && !guessAddrKey) {
            guessAddrKey = col;
          }
        });

        // Fallbacks if auto-guess fails
        if (!guessNameKey && rawColumns.length > 0) guessNameKey = rawColumns[0];
        if (!guessAddrKey && rawColumns.length > 1) guessAddrKey = rawColumns[1];
        if (!guessAddrKey) guessAddrKey = guessNameKey;

        setExcelMapping({
          show: true,
          columns: rawColumns,
          data: listData,
          nameKey: guessNameKey,
          addrKey: guessAddrKey
        });
      } catch (err) {
        alert("An error occurred while reading the file. Please export to standard Excel or CSV sheets.");
      }
    };
    reader.readAsArrayBuffer(file);
    // Reset file input so same file can be uploaded again
    if (event.target) event.target.value = "";
  };

  const applyMappingAndImport = () => {
    if (!excelMapping) return;
    const { data, nameKey, addrKey } = excelMapping;

    if (!nameKey || !addrKey) {
      alert("Please map both name and address columns to continue.");
      return;
    }

    const imported: School[] = data
      .map((row, index) => {
        const name = String(row[nameKey] || "").trim();
        const address = String(row[addrKey] || "").trim();

        if (!name && !address) return null;

        return {
          id: "import_" + Date.now().toString() + "_" + index,
          name: name || "Untitled Recipient",
          address: address || "No address details provided",
          selected: true
        };
      })
      .filter((s): s is School => s !== null);

    if (imported.length === 0) {
      alert("No valid entries found with the selected column mapping.");
      return;
    }

    setSchools((prev) => [...imported, ...prev]);
    setExcelMapping(null);
    setActiveTab("database");
    alert(`Successfully imported ${imported.length} new envelope items!`);
  };

  // Excel Export back to file download using 'xlsx'
  const handleExportDatabase = () => {
    try {
      const dataToExport = schools.map((s, index) => ({
        "Serial No": index + 1,
        "School Name": s.name,
        "Recipient Address": s.address,
        "Selected for Print": s.selected ? "Yes" : "No"
      }));

      const worksheet = XLSX.utils.json_to_sheet(dataToExport);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "Envelopes Directory");

      // Custom column widths
      worksheet["!cols"] = [
        { wch: 10 },
        { wch: 55 },
        { wch: 75 },
        { wch: 20 }
      ];

      XLSX.writeFile(workbook, `B2P_Envelopes_Directory_Database.xlsx`);
    } catch (err) {
      alert("Failed to export database. Please contact system support.");
    }
  };

  // Bulk set status
  const setSelectedStatusForAll = (status: boolean) => {
    setSchools((prev) => prev.map((s) => ({ ...s, selected: status })));
  };

  // Logo upload & removal handlers for custom sender brand profiles
  const handleLogoUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    if (!file.type.startsWith("image/")) {
      alert("Please upload a valid image file (PNG, JPG, SVG, etc.).");
      return;
    }

    const reader = new FileReader();
    reader.onload = (e) => {
      const dataUrl = e.target?.result as string;
      setSenders((prev) =>
        prev.map((s) => (s.id === activeSenderId ? { ...s, logoUrl: dataUrl } : s))
      );
    };
    reader.readAsDataURL(file);
  };

  const handleRemoveLogo = () => {
    setSenders((prev) =>
      prev.map((s) => {
        if (s.id === activeSenderId) {
          const updated = { ...s };
          delete updated.logoUrl;
          return updated;
        }
        return s;
      })
    );
  };

  const handleSelectPreset = (presetId: string) => {
    setSelectedPresetId(presetId);
    if (presetId === "custom") return;
    const preset = presets.find((p) => p.id === presetId);
    if (preset) {
      setToBlockTopPadding(preset.toBlockTopPadding);
      setToBlockXShift(preset.toBlockXShift);
      setToBlockWidth(preset.toBlockWidth);
    }
  };

  const handleSaveCurrentAsPreset = () => {
    triggerPrompt(
      "Save Printer Preset",
      "Enter a descriptive name for this alignment preset to easily load it in the future:",
      "Office Printer Tray 1",
      "Preset Name",
      (name) => {
        if (!name || !name.trim()) return;

        const newPreset: PrinterPreset = {
          id: "preset_user_" + Date.now().toString(),
          name: name.trim(),
          toBlockTopPadding,
          toBlockXShift,
          toBlockWidth
        };

        setPresets((prev) => [...prev, newPreset]);
        setSelectedPresetId(newPreset.id);
      }
    );
  };

  const handleDeletePreset = (id: string) => {
    const preset = presets.find((p) => p.id === id);
    if (!preset) return;
    if (preset.isBuiltIn) {
      triggerAlert("System Setting Locked", "This is a pre-configured built-in printer alignment standard and cannot be deleted.");
      return;
    }
    triggerConfirm(
      "Delete Alignment Preset",
      `Are you sure you want to delete the printer calibration preset "${preset.name}"? This action cannot be undone.`,
      () => {
        setPresets((prev) => prev.filter((p) => p.id !== id));
        handleSelectPreset(BUILTIN_PRESETS[0].id);
      }
    );
  };

  const handleRenamePreset = (id: string) => {
    const preset = presets.find((p) => p.id === id);
    if (!preset) return;
    if (preset.isBuiltIn) return;
    
    triggerPrompt(
      "Rename Printer Preset",
      "Assign a descriptive name for this alignment preset:",
      preset.name,
      "Preset Name",
      (newName) => {
        if (!newName || !newName.trim()) return;
        setPresets((prev) =>
          prev.map((p) => (p.id === id ? { ...p, name: newName.trim() } : p))
        );
      }
    );
  };

  const handlePrint = () => {
    if (selectedCount === 0) {
      alert("Please select at least 1 envelope database row to start printing.");
      return;
    }
    
    // Focus the active document window so browser prints correct target layout
    window.focus();
    
    try {
      window.print();
    } catch (err) {
      console.error("Print invocation failed:", err);
      alert(
        "Standard browser action blocked print in preview pane.\n\n" +
        "Please click the 'Open in New Tab' icon on the top right corner of the preview area or open the main URL in a direct tab to allow native print features."
      );
    }
  };

  // Quick select school for previewing inside left sidebar preview box
  const activePreviewSchool = useMemo(() => {
    if (selectedSchoolId) {
      const found = schools.find((s) => s.id === selectedSchoolId);
      if (found) return found;
    }
    // Fallback to first selected school
    const firstSelected = schools.find((s) => s.selected);
    if (firstSelected) return firstSelected;
    // Fallback to first school whatsoever
    return schools[0];
  }, [schools, selectedSchoolId]);

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col font-sans select-none antialiased text-slate-800">
      {/* Dynamic Header Toolbar (hidden in printed envelopes) */}
      <header className="bg-white border-b border-slate-200/80 shadow-sm no-print sticky top-0 z-50 transition-all duration-200">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-3.5 flex flex-col sm:flex-row items-center justify-between gap-4">
          
          {/* Logo and branding info */}
          <div className="flex items-center gap-3">
            {/* Elegant SVG-based Ligated modern B2P Logo matching design */}
            <div className="h-11 w-24 bg-slate-900 rounded-lg p-1.5 flex items-center justify-center border border-slate-800 shadow-inner">
              <svg viewBox="0 0 100 40" className="h-full w-full">
                <defs>
                  <linearGradient id="b2pGrad" x1="0%" y1="0%" x2="100%" y2="100%">
                    <stop offset="0%" stopColor="#EF4444" />
                    <stop offset="100%" stopColor="#B91C1C" />
                  </linearGradient>
                </defs>
                <path
                  d="M15 8 a6 6 0 0 1 6 6 v12 a6 6 0 0 1 -12 0 v-18 h4 v6 a6 6 0 0 1 2 0 a6 6 0 0 1 0 0 v6"
                  fill="none"
                  stroke="url(#b2pGrad)"
                  strokeWidth="3.5"
                  strokeLinecap="round"
                />
                <text x="36" y="26" fill="white" fontStyle="oblique" fontSize="15" fontWeight="900" letterSpacing="1">B2P</text>
                <text x="68" y="26" fill="#EF4444" fontSize="10" fontWeight="700">INT'L</text>
              </svg>
            </div>
            <div>
              <h1 id="b2p-app-title" className="text-sm font-extrabold tracking-tight text-slate-900 uppercase">
                Envelope Suite
              </h1>
              <p className="text-[10px] text-slate-400 font-mono">
                DL Premium 265mm &times; 112mm format calibrator
              </p>
            </div>
          </div>

          {/* Real-time print selection status badge / counter */}
          <div className="flex items-center gap-2">
            <div className="bg-slate-900/5 border border-slate-200/80 px-4 py-1.5 rounded-lg text-left flex items-center gap-3 shadow-sm">
              <div>
                <span className="block text-[9px] uppercase tracking-widest font-extrabold text-slate-400">
                  Select Status
                </span>
                <span className="text-xs text-slate-500 font-medium">To print queue</span>
              </div>
              <div className="bg-slate-900 text-white px-2.5 py-1 rounded-md flex items-center gap-1 font-mono font-bold text-xs shadow-inner">
                <span className="text-red-400">{selectedCount}</span>
                <span className="text-slate-500">/</span>
                <span className="text-slate-300">{schools.length}</span>
              </div>
            </div>
          </div>

          {/* Command controls & print triggers */}
          <div className="flex flex-wrap items-center gap-2 sm:gap-3">
            <button
              onClick={handlePrint}
              disabled={selectedCount === 0}
              className={`flex items-center gap-2 px-5 py-2 rounded-lg border text-xs font-bold shadow-sm transition-all cursor-pointer ${
                selectedCount > 0
                  ? "bg-red-600 hover:bg-red-700 text-white border-red-600 hover:shadow-md"
                  : "bg-slate-100 text-slate-400 border-slate-200 cursor-not-allowed"
              }`}
            >
              <Printer className="w-3.5 h-3.5" />
              <span>Print {selectedCount} Envelopes</span>
            </button>

            <button
              onClick={handleExportDatabase}
              className="flex items-center gap-2 px-4 py-2 bg-emerald-700 hover:bg-emerald-800 text-white rounded-lg border border-emerald-700 text-xs font-bold shadow-sm hover:shadow-md transition-all cursor-pointer"
              title="Export database back to Excel format with custom address coordinates"
            >
              <Download className="w-3.5 h-3.5" />
              <span>Excel Export</span>
            </button>
          </div>
        </div>

        {/* Workspace Tab navigation bar */}
        <div className="bg-slate-50/50 border-t border-slate-100">
          <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
            <nav className="flex space-x-1.5 py-2">
              <button
                onClick={() => setActiveTab("database")}
                className={`px-4 py-1.5 text-xs font-extrabold rounded-lg transition-all duration-150 cursor-pointer flex items-center gap-2 ${
                  activeTab === "database"
                    ? "bg-slate-900 text-white shadow-md shadow-slate-900/10"
                    : "text-slate-500 hover:text-slate-800 hover:bg-slate-200/50"
                }`}
              >
                <Grid className="w-3.5 h-3.5" />
                <span>Schools Directory</span>
              </button>

              <button
                onClick={() => setActiveTab("preview")}
                className={`px-4 py-1.5 text-xs font-extrabold rounded-lg transition-all duration-150 cursor-pointer flex items-center gap-2 ${
                  activeTab === "preview"
                    ? "bg-slate-900 text-white shadow-md shadow-slate-900/10"
                    : "text-slate-500 hover:text-slate-800 hover:bg-slate-200/50"
                }`}
              >
                <Eye className="w-3.5 h-3.5" />
                <span>Virtual Printed Gallery</span>
                {selectedCount > 0 && (
                  <span className="bg-red-500 text-white px-1.5 py-0.5 rounded-full text-[9px] font-mono font-bold leading-none">
                    {selectedCount}
                  </span>
                )}
              </button>

              <button
                onClick={() => setActiveTab("settings")}
                className={`px-4 py-1.5 text-xs font-extrabold rounded-lg transition-all duration-150 cursor-pointer flex items-center gap-2 ${
                  activeTab === "settings"
                    ? "bg-slate-900 text-white shadow-md shadow-slate-900/10"
                    : "text-slate-500 hover:text-slate-800 hover:bg-slate-200/50"
                }`}
              >
                <Sliders className="w-3.5 h-3.5" />
                <span>Company Brand Settings</span>
              </button>
            </nav>
          </div>
        </div>
      </header>

      {/* Interactive print guidance warning inside iframes */}
      {isInsideIframe && (
        <div className="bg-amber-50/90 backdrop-blur-sm border-b border-amber-200/80 p-3 no-print flex shrink-0 shadow-xs">
          <div className="max-w-7xl mx-auto w-full flex flex-col sm:flex-row items-center justify-between gap-3 text-xs text-amber-900 font-bold">
            <span className="flex items-center gap-2">
              <span className="h-2 w-2 rounded-full bg-amber-500 animate-ping shrink-0" />
              <span>Standard local print commands can be blocked by browsers inside secure nested previews.</span>
            </span>
            <button
              onClick={() => {
                window.open(window.location.href, "_blank");
              }}
              className="text-[10px] bg-amber-600 hover:bg-amber-700 text-white font-extrabold px-3 py-1.5 rounded-md transition-all shadow-sm cursor-pointer whitespace-nowrap uppercase tracking-wider"
            >
              Open direct workspace
            </button>
          </div>
        </div>
      )}

      {/* Main Multi-Column Workspace layout */}
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-6 flex-1 w-full flex flex-col lg:flex-row gap-6 no-print">
        
        {/* Left Control Panel / Real-time interactive branding view (hidden in print) */}
        <aside className="lg:w-80 flex flex-col gap-6 shrink-0 no-print">
          
          {/* Quick Upload / Import Action widget */}
          <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-5">
            <h3 className="text-[10px] font-extrabold uppercase tracking-widest text-slate-400 mb-3 flex items-center justify-between">
              <span>Import Sheet Database</span>
              <FileSpreadsheet className="w-3.5 h-3.5 text-slate-400" />
            </h3>
            
            <p className="text-xs text-slate-500 mb-3.5 leading-relaxed">
              Upload any school address roster (Excel <strong>.xlsx, .xls</strong>, or <strong>.csv</strong>). Drag directly onto this area:
            </p>

            <div
              onClick={() => fileInputRef.current?.click()}
              className="border-2 border-dashed border-slate-200 hover:border-red-500 bg-slate-50/50 hover:bg-red-50/10 rounded-xl p-5 text-center cursor-pointer transition-all duration-150 group"
            >
              <Upload className="w-6 h-6 mx-auto text-slate-400 group-hover:text-red-600 mb-2 transition-all duration-150 group-hover:scale-105" />
              <span className="block text-xs font-bold text-slate-700 group-hover:text-slate-900">
                Choose Excel or CSV File
              </span>
              <span className="block text-[10px] text-slate-400 mt-1">
                Col mapping auto-detected
              </span>
            </div>
            
            <input
              type="file"
              ref={fileInputRef}
              onChange={handleFileUpload}
              accept=".xlsx,.xls,.csv"
              className="hidden"
            />

            <div className="mt-3.5 flex gap-2 font-sans font-extrabold uppercase text-[10px] tracking-wider select-none">
              <button
                type="button"
                onClick={handleResetToPresets}
                className="flex-1 flex items-center justify-center gap-1.5 text-[10px] font-extrabold text-slate-600 hover:text-slate-900 py-2 border border-slate-200 hover:border-slate-300 rounded-lg bg-white transition-all cursor-pointer shadow-2xs hover:shadow-xs"
                title="Restore default 40 school records"
              >
                <RotateCcw className="w-3 h-3 text-slate-400" />
                <span>Reset 40</span>
              </button>
              <button
                type="button"
                onClick={handleDownloadTemplate}
                className="flex-1 flex items-center justify-center gap-1.5 text-[10px] font-extrabold text-red-600 hover:text-white py-2 border border-red-200 hover:border-red-600 rounded-lg bg-white hover:bg-red-600 transition-all cursor-pointer shadow-2xs hover:shadow-xs"
                title="Download standard Excel layout format"
              >
                <Download className="w-3 h-3" />
                <span>Template</span>
              </button>
            </div>
          </div>

          {/* Real-time scaled envelope element widget (Single example matching visual mockup) */}
          <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-5">
            <h3 className="text-[10px] font-extrabold uppercase tracking-widest text-slate-400 mb-3 flex items-center justify-between">
              <span>Preview Layout Scale</span>
              <div className="h-1.5 w-1.5 rounded-full bg-emerald-500 animate-pulse" />
            </h3>
            
            <div className="relative aspect-[265/112] w-full bg-slate-50 border border-slate-100 rounded-lg p-3 overflow-hidden shadow-xs flex flex-col justify-between select-none">
              <div className="flex justify-between items-start">
                
                {/* Standard Sender detail */}
                <div className="flex gap-1.5 items-start max-w-[65%] text-left">
                  {activeSender.logoUrl ? (
                    <div className="w-10 h-7 rounded bg-white shrink-0 p-0.5 border border-slate-200/60 flex items-center justify-center overflow-hidden">
                      <img src={activeSender.logoUrl} alt="Logo" className="max-h-full max-w-full object-contain pointer-events-none" referrerPolicy="no-referrer" />
                    </div>
                  ) : (
                    <div className="w-8 h-7 rounded bg-slate-900 shrink-0 p-0.5 flex items-center justify-center">
                      <span className="text-[6px] text-red-400 font-black font-sans leading-none">B2P</span>
                    </div>
                  )}
                  <div className="text-[5px] text-slate-500 scale-[0.85] origin-top-left font-sans leading-tight">
                    <span className="font-extrabold text-slate-900 uppercase block leading-none mb-0.5">{activeSender.name}</span>
                    {activeSender.address}
                  </div>
                </div>

                {/* Stamp block */}
                <div className="w-8 h-7 border border-dashed border-slate-300 flex items-center justify-center rounded bg-white shrink-0">
                  <span className="text-[3.5px] leading-tight text-center text-slate-400 font-extrabold uppercase tracking-wider scale-90">
                    STAMP
                  </span>
                </div>
              </div>

              {/* Recipient area preview mapping */}
              <div className="pl-10 text-left mb-1">
                <span className="text-[4.5px] font-extrabold text-slate-400 block mb-0.5 uppercase tracking-wider scale-90 origin-left">To:</span>
                {activePreviewSchool ? (
                  <>
                    <div className="text-[5.5px] font-black text-slate-900 line-clamp-1 truncate uppercase">
                      {activePreviewSchool.name}
                    </div>
                    <div className="text-[5px] text-slate-500 line-clamp-2 leading-snug whitespace-pre-wrap max-h-4 overflow-hidden">
                      {activePreviewSchool.address}
                    </div>
                  </>
                ) : (
                  <div className="text-[5px] italic text-slate-400">
                    No active school. Please check directory to select.
                  </div>
                )}
              </div>
            </div>

            <div className="mt-4 bg-slate-50 p-3 rounded-lg border border-slate-100">
              <span className="text-xs font-bold text-slate-700 block mb-1">
                Print Setup Alignment Rules:
              </span>
              <ul className="text-[11px] text-slate-500 space-y-1.5 list-disc pl-4 leading-normal">
                <li>Load envelopes in main or bypass manual tray.</li>
                <li>Set printer size template to <strong>265mm &times; 112mm</strong> or <strong>DL Format</strong>.</li>
                <li>Disable "Headers and Footers" in print preview settings.</li>
                <li>Set scale parameter specifically to <strong>100%</strong> (No padding fit-to-page scaling).</li>
              </ul>
            </div>
          </div>

          {/* Quick add manual row drawer panel */}
          <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-5">
            <h3 className="text-[10px] font-extrabold uppercase tracking-widest text-slate-400 mb-3.5 flex items-center gap-1.5">
              <Plus className="w-3.5 h-3.5 text-slate-400" />
              <span>Add Single Envelope</span>
            </h3>

            <div className="space-y-3">
              <div>
                <label className="block text-[10px] font-extrabold text-slate-500 uppercase tracking-wider mb-1">
                  Recipient Name / School
                </label>
                <input
                  type="text"
                  placeholder="e.g. St. Thomas Public School"
                  value={newSchoolName}
                  onChange={(e) => setNewSchoolName(e.target.value)}
                  className="w-full text-xs px-3 py-2 border border-slate-200 rounded-lg focus:outline-none focus:border-slate-900 focus:ring-2 focus:ring-slate-905/5 bg-slate-50/50 hover:bg-slate-50 font-medium"
                />
              </div>

              <div>
                <div className="flex justify-between items-center mb-1">
                  <label className="block text-[10px] font-extrabold text-slate-500 uppercase tracking-wider">
                    Complete Address Location
                  </label>
                  <button
                    type="button"
                    onClick={() => {
                      setNewSchoolAddr(autoFormatAddressText(newSchoolAddr));
                    }}
                    className="text-[9px] font-extrabold text-red-600 hover:text-white hover:bg-red-600 px-1.5 py-0.5 rounded border border-red-100 flex items-center gap-1 cursor-pointer transition-all uppercase"
                    title="Convert to Title Case and clean zip code formatting"
                  >
                    <Sparkles className="w-2 h-2" />
                    <span>Auto-Format</span>
                  </button>
                </div>
                <textarea
                  placeholder="e.g. Street, Region, Kerala - Pin Code"
                  rows={3}
                  value={newSchoolAddr}
                  onChange={(e) => setNewSchoolAddr(e.target.value)}
                  className="w-full text-xs px-3 py-2 border border-slate-200 rounded-lg focus:outline-none focus:border-slate-900 focus:ring-2 focus:ring-slate-905/5 bg-slate-50/50 hover:bg-slate-55 resize-none font-sans font-medium"
                />
              </div>

              <button
                onClick={handleAddNewSchool}
                disabled={!newSchoolName.trim() || !newSchoolAddr.trim()}
                className={`w-full py-2 rounded-lg text-xs font-bold flex items-center justify-center gap-1 transition-all cursor-pointer ${
                  newSchoolName.trim() && newSchoolAddr.trim()
                    ? "bg-slate-900 hover:bg-slate-800 text-white shadow-sm active:scale-98"
                    : "bg-slate-100 text-slate-400 cursor-not-allowed"
                }`}
              >
                <Plus className="w-3.5 h-3.5" />
                <span>Save to Local Database</span>
              </button>
            </div>
          </div>

        </aside>

        {/* Right workspace panel depending on selection */}
        <main className="flex-1 flex flex-col min-w-0">
          
          {/* Mapping Dialog Box */}
          {excelMapping && (
            <div className="mb-6 bg-orange-50 border-2 border-orange-200 rounded-xl p-5 shadow-sm text-slate-800">
              <div className="flex items-center justify-between border-b border-orange-200 pb-3 mb-4">
                <div className="flex items-center gap-2">
                  <FileSpreadsheet className="w-5 h-5 text-orange-600" />
                  <h3 className="text-sm font-extrabold text-orange-950">
                    Map Upload Columns from Spreadsheet
                  </h3>
                </div>
                <button
                  onClick={() => setExcelMapping(null)}
                  className="text-slate-400 hover:text-slate-600 cursor-pointer"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>

              <p className="text-xs text-slate-600 mb-4">
                Verify how your excel spreadsheet column titles line up to the envelope fields. Our pipeline matched them as follows:
              </p>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-4">
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1.5">
                    School / Recipient Name Column:
                  </label>
                  <select
                    value={excelMapping.nameKey}
                    onChange={(e) =>
                      setExcelMapping({ ...excelMapping, nameKey: e.target.value })
                    }
                    className="w-full text-xs bg-white p-2.5 border border-orange-300 rounded-lg focus:outline-none"
                  >
                    {excelMapping.columns.map((col) => (
                      <option key={col} value={col}>
                        {col}
                      </option>
                    ))}
                  </select>
                  <span className="text-[10px] text-slate-500 mt-1 block h-4 overflow-hidden text-ellipsis whitespace-nowrap">
                    Preview:{" "}{excelMapping.data[0]?.[excelMapping.nameKey] || "empty"}
                  </span>
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1.5">
                    Detailed Address Column:
                  </label>
                  <select
                    value={excelMapping.addrKey}
                    onChange={(e) =>
                      setExcelMapping({ ...excelMapping, addrKey: e.target.value })
                    }
                    className="w-full text-xs bg-white p-2.5 border border-orange-300 rounded-lg focus:outline-none"
                  >
                    {excelMapping.columns.map((col) => (
                      <option key={col} value={col}>
                        {col}
                      </option>
                    ))}
                  </select>
                  <span className="text-[10px] text-slate-500 mt-1 block h-4 overflow-hidden text-ellipsis whitespace-nowrap">
                    Preview:{" "}{excelMapping.data[0]?.[excelMapping.addrKey] || "empty"}
                  </span>
                </div>
              </div>

              <div className="flex justify-end gap-2.5">
                <button
                  onClick={() => setExcelMapping(null)}
                  className="px-4 py-2 text-xs font-bold text-slate-600 hover:bg-orange-100 rounded-lg cursor-pointer"
                >
                  Cancel
                </button>
                <button
                  onClick={applyMappingAndImport}
                  className="px-5 py-2 bg-orange-600 hover:bg-orange-700 text-white rounded-lg text-xs font-extrabold cursor-pointer transition-all shadow-sm"
                >
                  Confirm Columns & Import {excelMapping.data.length} records
                </button>
              </div>
            </div>
          )}

          {/* Tab Content 1: Database Directory */}
          {activeTab === "database" && (
            <div className="bg-white rounded-xl shadow-sm border border-slate-200/80 flex flex-col h-full flex-1 min-h-[450px]">
              
              {/* Directory Filter, Search, and Bulk Actions header */}
              <div className="p-4 border-b border-slate-200/60 flex flex-col md:flex-row gap-3 items-center justify-between bg-slate-50/30">
                
                {/* Search Bar Input */}
                <div className="relative w-full md:w-80">
                  <span className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none">
                    <Search className="h-3.5 w-3.5 text-slate-400" />
                  </span>
                  <input
                    type="text"
                    placeholder="Search schools or address details..."
                    value={searchQuery}
                    onChange={(e) => setSearchQuery(e.target.value)}
                    className="w-full pl-9 pr-4 py-2 text-xs border border-slate-200 rounded-lg focus:outline-none focus:border-slate-900 focus:ring-2 focus:ring-slate-900/5 bg-slate-100/40 hover:bg-slate-100/75 transition-all font-medium"
                  />
                  {searchQuery && (
                    <button
                      onClick={() => setSearchQuery("")}
                      className="absolute inset-y-0 right-0 pr-3 flex items-center text-slate-400 hover:text-slate-600 cursor-pointer"
                    >
                      <X className="h-3.5 w-3.5" />
                    </button>
                  )}
                </div>

                {/* Bulk status actions */}
                <div className="flex flex-wrap items-center gap-2 w-full md:w-auto justify-end">
                  <span className="text-[10px] font-extrabold uppercase tracking-wider text-slate-400 hidden sm:inline">
                    Bulk Actions:
                  </span>
                  
                  <button
                    onClick={() => setSelectedStatusForAll(true)}
                    className="px-3 py-1.5 bg-white hover:bg-slate-50 text-slate-700 border border-slate-200 rounded-lg text-[11px] font-extrabold cursor-pointer transition-all shadow-2xs hover:shadow-xs active:scale-98"
                  >
                    Select All ({schools.length})
                  </button>

                  <button
                    onClick={() => setSelectedStatusForAll(false)}
                    className="px-3 py-1.5 bg-white hover:bg-slate-50 text-slate-755 hover:text-slate-900 border border-slate-200 rounded-lg text-[11px] font-extrabold cursor-pointer transition-all shadow-2xs hover:shadow-xs active:scale-98"
                  >
                    Deselect All
                  </button>

                  <button
                    onClick={() => {
                      const count = schools.filter((s) => s.selected).length;
                      triggerConfirm(
                        "Bulk Delete Records",
                        `Are you absolutely sure you want to delete all ${count} selected recipient address records from this session? This action cannot be reversed.`,
                        () => {
                          setSchools((prev) => prev.filter((s) => !s.selected));
                        }
                      );
                    }}
                    disabled={selectedCount === 0}
                    className={`px-3 py-1.5 rounded-lg text-[11px] font-extrabold transition-all cursor-pointer border flex items-center gap-1 ${
                      selectedCount > 0
                        ? "bg-red-50 hover:bg-red-100/80 text-red-700 border-red-200 shadow-2xs"
                        : "bg-slate-50 text-slate-355 border-slate-100 cursor-not-allowed"
                    }`}
                  >
                    <Trash2 className="w-3 h-3" />
                    <span>Delete ({selectedCount})</span>
                  </button>
                </div>
              </div>

              {/* Data Table of schools */}
              <div className="flex-1 overflow-y-auto max-h-[600px] border-b border-slate-100">
                <table className="w-full text-left border-collapse">
                  <thead className="bg-slate-50/70 sticky top-0 md:relative z-10">
                    <tr className="border-b border-slate-100">
                      <th className="py-3 px-4 w-12 text-center">
                        <button
                          onClick={handleToggleSelectAll}
                          className="text-slate-400 hover:text-slate-600 transition-colors focus:outline-none cursor-pointer"
                          title="Toggle all search results"
                        >
                          {filteredSchools.length > 0 &&
                          filteredSchools.every((s) => s.selected) ? (
                            <CheckSquare className="w-4 h-4 text-red-600" />
                          ) : filteredSchools.some((s) => s.selected) ? (
                            <div className="w-4 h-4 border border-slate-300 rounded bg-red-100 flex items-center justify-center">
                              <span className="w-2 h-0.5 bg-red-600" />
                            </div>
                          ) : (
                            <Square className="w-4 h-4 text-slate-400" />
                          )}
                        </button>
                      </th>
                      <th className="py-3 px-3 text-xs font-extrabold uppercase tracking-widest text-slate-500 w-12 text-center">
                        #
                      </th>
                      <th className="py-3 px-4 text-xs font-extrabold uppercase tracking-widest text-slate-500">
                        School / Institution Name
                      </th>
                      <th className="py-3 px-4 text-xs font-extrabold uppercase tracking-widest text-slate-500">
                        Recipient Address
                      </th>
                      <th className="py-3 px-4 text-xs font-extrabold uppercase tracking-widest text-slate-500 w-28 text-center">
                        Actions
                      </th>
                    </tr>
                  </thead>
                  
                  <tbody className="divide-y divide-slate-100 text-slate-700">
                    {filteredSchools.length > 0 ? (
                      filteredSchools.map((school, index) => (
                        <tr
                          key={school.id}
                          className={`hover:bg-slate-50/50 transition-all ${
                            school.selected ? "bg-red-50/20" : ""
                          }`}
                        >
                          <td className="py-3.5 px-4 text-center">
                            <button
                              onClick={() => handleToggleSingleSelect(school.id)}
                              className="text-slate-400 hover:text-red-600 focus:outline-none cursor-pointer"
                            >
                              {school.selected ? (
                                <CheckSquare className="w-4 h-4 text-red-600" id={`school-checkbox-${school.id}`} />
                              ) : (
                                <Square className="w-4 h-4" id={`school-checkbox-${school.id}`} />
                              )}
                            </button>
                          </td>
                          <td className="py-3.5 px-3 text-center text-xs font-mono text-slate-400">
                            {index + 1}
                          </td>
                          <td className="py-3.5 px-4 max-w-[240px] truncate-cell">
                            <div className="text-xs font-bold text-slate-900 leading-tight">
                              {school.name}
                            </div>
                          </td>
                          <td className="py-3.5 px-4 max-w-sm">
                            <div className="text-xs text-slate-500 leading-relaxed font-sans whitespace-pre-line truncate max-h-12 overflow-hidden hover:max-h-none transition-all">
                              {school.address}
                            </div>
                          </td>
                          <td className="py-3.5 px-4 text-center">
                            <div className="flex items-center justify-center gap-1.5">
                              <button
                                onClick={() => handleStartEdit(school)}
                                className="p-1 px-1.5 text-slate-500 hover:text-slate-800 hover:bg-slate-100 rounded transition-all cursor-pointer text-xs flex items-center gap-0.5"
                                title="Edit Address Coordinates"
                              >
                                <Edit2 className="w-3 h-3" />
                                <span>Edit</span>
                              </button>
                              
                              <button
                                onClick={() => handleDeleteSchool(school.id)}
                                className="p-1 px-1.5 text-slate-400 hover:text-slate-800 hover:bg-rose-50 rounded transition-all cursor-pointer text-xs"
                                title="Delete Record"
                              >
                                <X className="w-3.5 h-3.5" />
                              </button>
                            </div>
                          </td>
                        </tr>
                      ))
                    ) : (
                      <tr>
                        <td colSpan={5} className="py-12 text-center text-slate-400">
                          <Building className="w-12 h-12 mx-auto text-slate-200 mb-2" />
                          <p className="text-xs font-bold text-slate-600">No schools match search query</p>
                          <p className="text-[11px] text-slate-400 mt-0.5">Toggle filter flags or import another Excel sheet</p>
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>

              {/* Toolbar metadata stats */}
              <div className="p-4 bg-slate-50 border-t border-slate-100 rounded-b-xl flex flex-col sm:flex-row items-center justify-between text-[11px] text-slate-400 font-mono gap-2">
                <div>
                  <span>Total schools registered: <strong>{schools.length}</strong></span>
                  <span className="mx-2 text-slate-300">|</span>
                  <span>Filtered visualizers: <strong>{filteredSchools.length}</strong></span>
                </div>
                <div className="flex items-center gap-1 text-slate-400">
                  <Clock className="w-3 h-3 text-slate-400" />
                  <span>Envelopes formatted &amp; updated to latest standards</span>
                </div>
              </div>
            </div>
          )}

          {/* Tab Content 2: Print Preview Stage (Each at exact DL 265mm x 112mm printed dimension) */}
          {activeTab === "preview" && (
            <div className="flex flex-col lg:flex-row gap-6 items-start">
              
              {/* Left Fine-Tuning controls sidebar (hidden when printing) */}
              {showCalibrationSidebar && (
                <div className="w-full lg:w-80 shrink-0 space-y-5 no-print animate-fade-in">
                  <div className="bg-white rounded-xl shadow-xs border border-slate-200 p-5 space-y-5">
                  <div className="border-b border-slate-200/60 pb-3 mb-1">
                    <h3 className="text-xs font-black uppercase tracking-widest text-slate-800 flex items-center gap-2">
                      <Sliders className="w-3.5 h-3.5 text-red-650" />
                      <span>Styles &amp; Typography Calibration</span>
                    </h3>
                  </div>

                  {/* Selector: Custom Font family */}
                  <div className="space-y-2">
                    <label className="block text-[10px] font-extrabold uppercase tracking-wider text-slate-500">
                      Unified Envelope Font:
                    </label>
                    <select
                      value={envelopeFont}
                      onChange={(e) => setEnvelopeFont(e.target.value)}
                      className="w-full text-xs bg-white p-2.5 border border-slate-200 rounded-lg focus:outline-none focus:border-slate-900 font-bold text-slate-705 shadow-2xs cursor-pointer"
                    >
                      <option value="Arial, Helvetica, sans-serif">Arial / Helvetica (Default Swiss)</option>
                      <option value='"Plus Jakarta Sans", sans-serif'>Plus Jakarta Sans</option>
                      <option value='"Inter", sans-serif'>Inter</option>
                      <option value='"Georgia", serif'>Georgia (Elegant Serif)</option>
                      <option value='"Courier New", monospace'>Courier New (Raw Polytype)</option>
                    </select>
                    <span className="text-[10px] text-slate-400 block leading-tight font-sans">
                      Ensures every text element inside the preview and physical print layout shares the exact same custom typeface.
                    </span>
                  </div>

                  {/* Printer Presets Block */}
                  <div className="space-y-3 pt-1 border-t border-slate-100">
                    <div className="flex items-center justify-between">
                      <label className="block text-[10px] font-extrabold uppercase tracking-wider text-slate-500">
                        Printer Alignment Presets:
                      </label>
                      {selectedPresetId === "custom" && (
                        <span className="text-[8px] font-mono font-bold text-amber-600 bg-amber-50 px-1.5 py-0.5 rounded border border-amber-100 uppercase animate-pulse">
                          Unsaved Changes
                        </span>
                      )}
                    </div>

                    <div className="flex gap-2 items-center">
                      <select
                        value={selectedPresetId}
                        onChange={(e) => handleSelectPreset(e.target.value)}
                        className="flex-1 text-xs bg-white p-2.5 border border-slate-200 rounded-lg focus:outline-none focus:border-slate-900 font-bold shadow-2xs text-slate-700 cursor-pointer"
                      >
                        <option value="custom" disabled={selectedPresetId !== "custom"}>
                          {selectedPresetId === "custom" ? "Custom Adjustments (Sliders)" : "-- Select Preset --"}
                        </option>
                        {presets.map((p) => (
                          <option key={p.id} value={p.id}>
                            {p.name} {p.isBuiltIn ? "(Built-in)" : "(Saved)"}
                          </option>
                        ))}
                      </select>

                      {/* Delete / Rename Custom Preset Button */}
                      {selectedPresetId !== "custom" && !presets.find((p) => p.id === selectedPresetId)?.isBuiltIn && (
                        <div className="flex gap-1 shrink-0">
                          <button
                            onClick={() => handleRenamePreset(selectedPresetId)}
                            className="p-2.5 bg-white hover:bg-slate-50 text-slate-700 rounded-lg cursor-pointer transition-all border border-slate-205 shadow-2xs"
                            title="Rename selected preset"
                          >
                            <Edit2 className="w-3.5 h-3.5" />
                          </button>
                          <button
                            onClick={() => handleDeletePreset(selectedPresetId)}
                            className="p-2.5 bg-red-50 hover:bg-red-650 hover:text-white text-red-650 rounded-lg cursor-pointer transition-all border border-red-200 shadow-2xs"
                            title="Delete selected preset"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>
                        </div>
                      )}
                    </div>

                    {/* Quick Alignment Description for current preset */}
                    <div className="text-[10px] text-slate-400 bg-white p-2.5 rounded-lg border border-slate-200/50 space-y-1 font-sans">
                      <div className="flex justify-between">
                        <span>Vertical Position:</span>
                        <span className="font-mono font-bold text-slate-600">{toBlockTopPadding}mm</span>
                      </div>
                      <div className="flex justify-between">
                        <span>Horizontal Shift:</span>
                        <span className="font-mono font-bold text-slate-600">{toBlockXShift}mm</span>
                      </div>
                      <div className="flex justify-between">
                        <span>Max Width:</span>
                        <span className="font-mono font-bold text-slate-600">{toBlockWidth}mm</span>
                      </div>
                    </div>

                    {/* Action to save current slider settings as new preset */}
                    <button
                      onClick={handleSaveCurrentAsPreset}
                      className="w-full text-xs text-red-600 hover:text-white font-extrabold bg-red-50 hover:bg-red-600 p-2.5 rounded-lg border border-red-200 hover:border-red-600 cursor-pointer flex items-center justify-center gap-1.5 transition-all shadow-2xs active:scale-98"
                    >
                      <Plus className="w-3.5 h-3.5" />
                      <span>Save Current Alignments as Preset</span>
                    </button>
                  </div>

                  {/* TO Recipient Address Block Calibration Widget */}
                  <div className="bg-slate-50/50 p-4 border border-slate-200/60 rounded-xl space-y-4 shadow-2xs border-t border-slate-100">
                    <label className="block text-[10px] font-black uppercase tracking-wider text-slate-500 mb-1">
                      To Address Area Alignment &amp; Fine-Tuning:
                    </label>

                    {/* Slider: Top Padding (Vertical Offset) */}
                    <div>
                      <div className="flex justify-between items-center mb-1 font-sans">
                        <span className="text-xs font-bold text-slate-700">
                          Vertical Position (Top Space):
                        </span>
                        <span className="text-xs font-mono font-bold text-red-600 bg-white border border-red-100 px-1.5 py-0.5 rounded shadow-3xs">
                          {toBlockTopPadding}mm
                        </span>
                      </div>
                      <input
                        type="range"
                        min="5"
                        max="40"
                        step="1"
                        value={toBlockTopPadding}
                        onChange={(e) => setToBlockTopPadding(Number(e.target.value))}
                        className="w-full accent-red-600 h-1.5 bg-slate-200 rounded-lg cursor-pointer"
                      />
                    </div>

                    {/* Slider: Horizontal Shift */}
                    <div>
                      <div className="flex justify-between items-center mb-1 font-sans">
                        <span className="text-xs font-bold text-slate-700">
                          Horizontal Shift (Offset X):
                        </span>
                        <span className="text-xs font-mono font-bold text-red-600 bg-white border border-red-100 px-1.5 py-0.5 rounded shadow-3xs">
                          {toBlockXShift > 0 ? `+${toBlockXShift}` : toBlockXShift}mm
                        </span>
                      </div>
                      <input
                        type="range"
                        min="-30"
                        max="30"
                        step="1"
                        value={toBlockXShift}
                        onChange={(e) => setToBlockXShift(Number(e.target.value))}
                        className="w-full accent-red-600 h-1.5 bg-slate-200 rounded-lg cursor-pointer"
                      />
                    </div>

                    {/* Slider: Max Width */}
                    <div>
                      <div className="flex justify-between items-center mb-1 font-sans">
                        <span className="text-xs font-bold text-slate-700">
                          Max Print block Width:
                        </span>
                        <span className="text-xs font-mono font-bold text-red-600 bg-white border border-red-100 px-1.5 py-0.5 rounded shadow-3xs">
                          {toBlockWidth}mm
                        </span>
                      </div>
                      <input
                        type="range"
                        min="100"
                        max="160"
                        step="2"
                        value={toBlockWidth}
                        onChange={(e) => setToBlockWidth(Number(e.target.value))}
                        className="w-full accent-red-600 h-1.5 bg-slate-200 rounded-lg cursor-pointer"
                      />
                    </div>
                  </div>

                  {/* Recipient Letter Spacing */}
                  <div className="border-t border-slate-100 pt-4 space-y-1">
                    <div className="flex justify-between items-center mb-1">
                      <label className="text-xs font-bold text-slate-700">
                        Recipient Address Letter Spacing:
                      </label>
                      <span className="text-xs font-mono font-bold text-red-600 bg-red-50 px-2 py-0.5 rounded shadow-3xs">
                        {letterSpacing}px
                      </span>
                    </div>
                    <input
                      type="range"
                      min="0"
                      max="10"
                      step="1"
                      value={letterSpacing}
                      onChange={(e) => setLetterSpacing(Number(e.target.value))}
                      className="w-full accent-red-600 h-1.5 bg-slate-100 rounded-lg cursor-pointer"
                    />
                    <span className="text-[10px] text-slate-400 mt-1 block leading-normal font-sans">
                      High letter-spacing matches the luxury/international style demonstrated in target mockup printout.
                    </span>
                  </div>

                  {/* Sender Letter Spacing */}
                  <div className="border-t border-slate-100 pt-4 space-y-1">
                    <div className="flex justify-between items-center mb-1">
                      <label className="text-xs font-bold text-slate-700">
                        Sender Letter Spacing:
                      </label>
                      <span className="text-xs font-mono font-bold text-red-600 bg-red-50 px-2 py-0.5 rounded shadow-3xs">
                        {fromLetterSpacing}px
                      </span>
                    </div>
                    <input
                      type="range"
                      min="0"
                      max="8"
                      step="1"
                      value={fromLetterSpacing}
                      onChange={(e) => setFromLetterSpacing(Number(e.target.value))}
                      className="w-full accent-red-600 h-1.5 bg-slate-100 rounded-lg cursor-pointer"
                    />
                  </div>

                </div>
              </div>
              )}

              {/* Right Workspace Stage Column: Interactive printed envelopes preview list */}
              <div className="flex-1 min-w-0 w-full space-y-6">
                
                {/* Interactive Workspace Control Toolbar */}
                <div className="bg-white rounded-xl shadow-xs border border-slate-200 p-4 flex flex-col md:flex-row items-center justify-between gap-4 no-print select-none">
                  <div className="flex items-center gap-3">
                    <button
                      onClick={() => setShowCalibrationSidebar(!showCalibrationSidebar)}
                      className="flex items-center gap-2 bg-slate-50 hover:bg-slate-100 text-slate-750 hover:text-slate-900 text-xs font-bold uppercase tracking-wider px-3.5 py-2.5 rounded-lg border border-slate-200 transition-all cursor-pointer shadow-3xs"
                      title={showCalibrationSidebar ? "Hide Calibration Sliders" : "Show Calibration Sliders"}
                    >
                      {showCalibrationSidebar ? <Eye className="w-4 h-4 text-slate-500" /> : <Sliders className="w-4 h-4 text-red-650" />}
                      <span>{showCalibrationSidebar ? "Maximize Screen Layout" : "Calibrations Panel"}</span>
                    </button>
                    
                    <span className="text-xs text-slate-300 font-bold hidden sm:inline">|</span>
                    
                    <div className="hidden sm:flex items-center gap-1.5 text-xs text-slate-500 font-bold">
                      <span className="bg-red-50 text-red-650 px-2 py-0.5 rounded text-[10px] font-mono tracking-wider font-extrabold border border-red-100 uppercase">
                        DL Formatted Model
                      </span>
                    </div>
                  </div>

                  {/* Sizing & zoom engine */}
                  <div className="flex items-center gap-4 w-full md:w-auto justify-end">
                    <div className="flex items-center gap-2 w-full max-w-xs md:w-44 font-sans">
                      <span className="text-[10px] font-extrabold uppercase tracking-widest text-slate-400 shrink-0 select-none">
                        Scale Preview:
                      </span>
                      <input
                        type="range"
                        min="30"
                        max="100"
                        step="5"
                        value={previewZoom}
                        onChange={(e) => setPreviewZoom(Number(e.target.value))}
                        className="w-full accent-red-600 h-1.5 bg-slate-100 rounded-lg cursor-pointer"
                      />
                      <span className="text-xs font-mono font-black text-red-650 bg-red-50 px-2 py-0.5 rounded shadow-3xs border border-red-100/50 min-w-[44px] text-center">
                        {previewZoom}%
                      </span>
                    </div>

                    <div className="shrink-0 flex gap-1.5 border-l border-slate-200 pl-3">
                      <button
                        onClick={() => setPreviewZoom(50)}
                        className={`text-[10px] font-bold px-2 py-1.5 rounded transition-all cursor-pointer border ${
                          previewZoom === 50
                            ? "bg-red-600 border-red-600 text-white"
                            : "bg-white hover:bg-slate-50 text-slate-600 border-slate-200 shadow-3xs"
                        }`}
                      >
                        50%
                      </button>
                      <button
                        onClick={() => setPreviewZoom(75)}
                        className={`text-[10px] font-bold px-2 py-1.5 rounded transition-all cursor-pointer border ${
                          previewZoom === 75
                            ? "bg-red-600 border-red-600 text-white"
                            : "bg-white hover:bg-slate-50 text-slate-600 border-slate-200 shadow-3xs"
                        }`}
                      >
                        75% (Ideal)
                      </button>
                      <button
                        onClick={() => setPreviewZoom(100)}
                        className={`text-[10px] font-bold px-2 py-1.5 rounded transition-all cursor-pointer border ${
                          previewZoom === 100
                            ? "bg-red-600 border-red-600 text-white"
                            : "bg-white hover:bg-slate-50 text-slate-600 border-slate-200 shadow-3xs"
                        }`}
                      >
                        100%
                      </button>
                    </div>
                  </div>
                </div>
                
                {/* Toolbar align message panel */}
                <div className="bg-amber-50 border border-amber-200/80 rounded-xl p-4 flex items-start gap-3 text-amber-900 shadow-sm no-print">
                  <Info className="w-5 h-5 text-amber-600 shrink-0 mt-0.5" />
                  <div className="text-xs space-y-1.5">
                    <p className="font-bold">
                      Interactive Physical Print Calibration (232mm &times; 112mm Active Workspace / DL Format Standard)
                    </p>
                    <p className="text-amber-800 leading-relaxed font-sans">
                      This gallery draws exact page-break limits for your print stream. When you click <strong>"Print Envelopes"</strong>, a high fidelity CSS media engine replaces web styles to yield exactly 1 school envelope per physical sheet. We highly advise a dry-run test with standard A4 paper before running on actual premium cardstocks!
                    </p>
                  </div>
                </div>

                 {/* Envelope list scroll mockup preview */}
                <div className="bg-slate-900/[0.03] p-6 lg:p-10 rounded-2xl border border-slate-200/60 shadow-inner space-y-12 max-w-full overflow-x-auto">
                {schools.filter((s) => s.selected).length > 0 ? (
                  schools
                    .filter((s) => s.selected)
                    .map((school, i) => {
                      const scaleValue = previewZoom / 100;
                      const isFirefox = typeof navigator !== "undefined" && navigator.userAgent.toLowerCase().includes("firefox");
                      
                      return (
                        <div
                          key={school.id}
                          className="bg-stone-50/95 rounded-lg shadow-[0_12px_30px_rgba(0,0,0,0.06)] border border-stone-200 mx-auto transition-transform hover:-translate-y-0.5 hover:shadow-[0_20px_40px_rgba(0,0,0,0.09)] relative overflow-hidden shrink-0 flex flex-col justify-between origin-top"
                          style={{
                            width: "265mm",
                            height: "112mm",
                            padding: "5mm 8mm 6mm 4.5mm",
                            aspectRatio: "265/112",
                            fontFamily: envelopeFont,
                            zoom: isFirefox ? undefined : scaleValue,
                            transform: isFirefox ? `scale(${scaleValue})` : undefined,
                            transformOrigin: isFirefox ? "top center" : undefined,
                            marginBottom: isFirefox ? `calc(112mm * (${scaleValue} - 1))` : undefined,
                          }}
                        >
                        {/* Interactive Page numbering identifier inside Preview (ignored in Print block via class standard) */}
                        <div className="absolute top-2 right-40 bg-slate-100 border border-slate-200 px-2 py-0.5 rounded text-[9px] font-extrabold font-mono text-slate-500 tracking-wider no-print shadow-sm z-10 flex items-center gap-1 uppercase select-none">
                          <Building className="w-2.5 h-2.5" />
                          <span>Envelope Unit {i + 1}</span>
                        </div>

                        {/* Top Sender & Stamp row */}
                        <div className="flex justify-between items-start" style={{ height: "27mm" }}>
                          
                          {/* Left sender block */}
                          <div className="flex items-start">
                            {activeSender.logoUrl ? (
                              <div
                                className="mr-3 shrink-0 border border-slate-100 bg-white rounded-md p-1 flex items-center justify-center overflow-hidden"
                                style={{
                                  width: "42mm",
                                  height: "24mm",
                                  minWidth: "42mm",
                                  minHeight: "24mm"
                                }}
                              >
                                <img
                                  src={activeSender.logoUrl}
                                  alt="Brand Logo"
                                  style={{
                                    maxWidth: "100%",
                                    maxHeight: "100%",
                                    width: "auto",
                                    height: "auto",
                                    objectFit: "contain"
                                  }}
                                  referrerPolicy="no-referrer"
                                />
                              </div>
                            ) : (
                              /* B2P styled beautiful Logo branding */
                              <svg
                                viewBox="0 0 100 40"
                                className="mr-3 shrink-0 block"
                                style={{
                                  width: "42mm",
                                  height: "24mm",
                                  minWidth: "42mm",
                                  minHeight: "24mm",
                                  fontFamily: envelopeFont
                                }}
                              >
                                <defs>
                                  <linearGradient id="logoGrad" x1="0%" y1="0%" x2="100%" y2="100%">
                                    <stop offset="0%" stopColor="#EA580C" />
                                    <stop offset="100%" stopColor="#DC2626" />
                                  </linearGradient>
                                </defs>
                                <path
                                  d="M12 4 a8 8 0 0 1 8 8 v14 a8 8 0 0 1 -16 0 v-20 h5 v6 a8 8 0 0 1 3 0 a8 8 0 0 1 0 0 v8"
                                  fill="none"
                                  stroke="url(#logoGrad)"
                                  strokeWidth="4.5"
                                  strokeLinecap="round"
                                />
                                <text x="36" y="24" fill="#0F172A" fontStyle="oblique" fontSize="17" fontWeight="900" letterSpacing="1">B2P</text>
                                <text x="36" y="34" fill="#64748B" fontSize="8" fontWeight="600" letterSpacing="0.5">International</text>
                              </svg>
                            )}

                            {/* Detailed Sender Address */}
                            <div className="from-details text-left shrink-0" style={{ marginLeft: "7mm", paddingTop: "1.8mm" }}>
                              <span
                                className="from-label block text-[#090909] font-extrabold uppercase"
                                style={{
                                  fontSize: "12pt",
                                  lineHeight: "1",
                                  letterSpacing: `${fromLetterSpacing}px`,
                                  marginBottom: "3mm",
                                  fontFamily: envelopeFont
                                }}
                              >
                                FROM:
                              </span>
                              
                              <div
                                className="from-name text-[#151515] font-extrabold"
                                style={{
                                  fontSize: "11.5pt",
                                  lineHeight: "1.1",
                                  letterSpacing: `${fromLetterSpacing}px`,
                                  marginBottom: "2.2mm",
                                  fontFamily: envelopeFont
                                }}
                              >
                                {activeSender.name}
                              </div>

                              <div
                                className="from-addr text-[#4a4a4a] font-normal"
                                style={{
                                  fontSize: "9.5pt",
                                  lineHeight: "1.55",
                                  letterSpacing: "0px",
                                  fontFamily: envelopeFont
                                }}
                              >
                                {activeSender.address.split("\n").map((line, idx) => (
                                  <span key={idx} className="block whitespace-nowrap">
                                    {line}
                                  </span>
                                ))}
                              </div>
                            </div>
                          </div>

                          {/* Stamp box widget area */}
                          <div
                            className={`flex items-center justify-center shrink-0 ${
                              enableStampBorder ? "border border-slate-300 rounded" : ""
                            }`}
                            style={{
                              width: "29.5mm",
                              height: "25mm",
                              marginTop: "1.5mm"
                            }}
                          >
                            <span
                              className="text-slate-600 text-center uppercase whitespace-pre-wrap leading-relaxed tracking-wider select-none"
                              style={{
                                fontSize: "6.5pt",
                                fontWeight: 700,
                                letterSpacing: "1.5px",
                                fontFamily: envelopeFont
                              }}
                            >
                              {stampText}
                            </span>
                          </div>
                        </div>

                        {/* Mid-envelope spacer block */}
                        <div style={{ height: "4.7mm" }} />

                        {/* Lower To Recipient Block exactly sized */}
                        <div
                          className="to-block flex flex-col items-start text-left"
                          style={{
                            width: `${toBlockWidth}mm`,
                            marginLeft: "auto",
                            paddingTop: `${toBlockTopPadding}mm`, // custom offset margin
                            transform: `translateX(${toBlockXShift}mm)`
                          }}
                        >
                          {showToPrefix && (
                            <span
                              className="to-label block font-extrabold text-[#050505] text-left"
                              style={{
                                fontSize: "14pt",
                                lineHeight: "1",
                                letterSpacing: `${letterSpacing}px`,
                                marginBottom: "6mm",
                                fontFamily: envelopeFont
                              }}
                            >
                              To:
                            </span>
                          )}

                          <div
                            className="to-name text-slate-900 font-extrabold text-left capitalize tracking-normal"
                            style={{
                              fontSize: "13.5pt",
                              lineHeight: "1.45",
                              letterSpacing: `${letterSpacing}px`,
                              marginBottom: "1.5mm",
                              whiteSpace: "normal",
                              fontFamily: envelopeFont
                            }}
                          >
                            {school.name}
                          </div>

                          <div
                            className="to-addr text-slate-800 font-medium text-left tracking-normal whitespace-pre-line leading-relaxed"
                            style={{
                              fontSize: "13pt",
                              lineHeight: "1.45",
                              letterSpacing: `${letterSpacing}px`,
                              fontFamily: envelopeFont
                            }}
                          >
                            {school.address}
                          </div>
                        </div>
                      </div>
                    );
                  })
                ) : (
                  <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-12 text-center text-slate-400 max-w-md mx-auto">
                    <Printer className="w-16 h-16 mx-auto text-slate-200 mb-2" />
                    <p className="text-sm font-extrabold text-slate-700">No Printing Candidates Registered</p>
                    <p className="text-xs text-slate-400 mt-1.5 leading-relaxed">
                      Go to the <strong>"Schools Directory"</strong> tab and checkmark the desired school boxes to populate the preview list!
                    </p>
                  </div>
                )}
              </div>
            </div>
          </div>
        )}

          {/* Tab Content 3: Branding & Alignment Adjustments Settings */}
          {activeTab === "settings" && (
            <div className="bg-white rounded-xl shadow-xs border border-slate-200/80 p-6 space-y-6">
              
              <div className="border-b border-slate-200/60 pb-3.5 mb-4">
                <h3 className="text-sm font-extrabold text-slate-900 flex items-center gap-2 uppercase tracking-tight">
                  <Sliders className="w-4 h-4 text-slate-800" />
                  <span>Branding &amp; Alignment Settings</span>
                </h3>
                <p className="text-xs text-slate-400 mt-0.5 leading-relaxed">
                  Calibrate company name labels, stamp values, layout positions, and custom spacings to match your business letter envelopes.
                </p>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                
                {/* Form fields: Sender Parameters */}
                <div className="space-y-5">
                  <h4 className="text-[10px] font-extrabold text-slate-400 uppercase tracking-widest flex items-center gap-1.5">
                    <Building className="w-3.5 h-3.5 text-slate-400" />
                    <span>Corporate Sender Profiles</span>
                  </h4>

                  {/* Profile selector dropdown and ADD NEW button */}
                  <div className="bg-slate-50/50 p-4 border border-slate-200/60 rounded-xl space-y-3">
                    <div className="flex items-center justify-between">
                      <label className="text-[10px] font-extrabold uppercase tracking-wider text-slate-500 block">
                        Select Active Corporate Unit:
                      </label>
                      <button
                        onClick={() => {
                          triggerDoublePrompt(
                            "Add Corporate Sender Profile",
                            "Create a new company sender card layout by defining its name and full physical address below:",
                            "GLOBAL LOGISTIC ENTERPRISE",
                            "Company / Brand Name",
                            "123 Business Avenue, Suite B\nThrissur, Kerala, India - 680001",
                            "Complete Dispatch Address Location",
                            (newCompanyName, newCompanyAddress) => {
                              if (!newCompanyName || !newCompanyName.trim() || !newCompanyAddress || !newCompanyAddress.trim()) return;
                              const newId = "company_" + Date.now().toString();
                              const newProfile: SenderInfo = {
                                id: newId,
                                name: newCompanyName.trim(),
                                address: newCompanyAddress.trim()
                              };
                              setSenders((prev) => [...prev, newProfile]);
                              setActiveSenderId(newId);
                            }
                          );
                        }}
                        className="text-[9px] text-red-600 hover:text-white font-extrabold flex items-center gap-1 cursor-pointer bg-red-50 hover:bg-red-600 p-1 py-1.5 px-3 rounded-lg border border-red-200 transition-all uppercase tracking-wider shadow-2xs hover:shadow-xs"
                      >
                        <Plus className="w-3 h-3" />
                        <span>Add Company</span>
                      </button>
                    </div>

                    <div className="flex items-center gap-2">
                       <select
                        value={activeSenderId}
                        onChange={(e) => setActiveSenderId(e.target.value)}
                        className="flex-1 text-xs bg-white p-2.5 border border-slate-205 rounded-lg focus:outline-none focus:border-slate-900 font-bold text-slate-705 shadow-2xs"
                      >
                        {senders.map((profile) => (
                          <option key={profile.id} value={profile.id}>
                            {profile.name}
                          </option>
                        ))}
                      </select>

                      {senders.length > 1 && (
                        <button
                          onClick={() => {
                            triggerConfirm(
                              "Remove Company Profile",
                              `Are you absolutely sure you want to permanently delete the "${activeSender.name}" corporate sender profile? This action cannot be reversed.`,
                              () => {
                                const remaining = senders.filter((s) => s.id !== activeSenderId);
                                setSenders(remaining);
                                setActiveSenderId(remaining[0].id);
                              }
                            );
                          }}
                          className="p-2.5 bg-red-50 border border-red-200 hover:bg-red-600 hover:text-white text-red-600 rounded-lg cursor-pointer transition-all active:scale-95"
                          title="Delete current company profile"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </button>
                      )}
                    </div>
                  </div>

                  {/* Upload custom logo for active sender */}
                  <div className="p-4 bg-slate-50/50 border border-slate-200/60 rounded-xl space-y-3">
                    <label className="block text-[10px] font-extrabold uppercase tracking-wider text-slate-500">
                      Corporate Logo Upload:
                    </label>

                    {activeSender.logoUrl ? (
                      <div className="flex items-center gap-4">
                        <div className="h-16 w-32 bg-white border border-slate-200 rounded-lg p-1.5 flex items-center justify-center overflow-hidden shadow-2xs">
                          <img
                            src={activeSender.logoUrl}
                            alt="Custom uploaded company logo preview"
                            className="max-h-full max-w-full object-contain pointer-events-none"
                            referrerPolicy="no-referrer"
                          />
                        </div>
                        <div className="space-y-1.5 flex-1 text-left">
                          <button
                            onClick={handleRemoveLogo}
                            className="bg-white border border-slate-200 hover:bg-red-50 hover:text-red-600 text-slate-600 px-3 py-1.5 text-xs font-extrabold rounded-lg cursor-pointer transition-all flex items-center gap-1 shadow-2xs hover:shadow-xs active:scale-98"
                          >
                            <Trash2 className="w-3.5 h-3.5 text-red-500" />
                            <span>Remove Uploaded Logo</span>
                          </button>
                          <span className="text-[10px] text-slate-400 block leading-tight">
                            Reverts to fallback standard high contrast B2P visual brand logo.
                          </span>
                        </div>
                      </div>
                    ) : (
                      <div className="space-y-2">
                        <div
                          onClick={() => logoInputRef.current?.click()}
                          className="border-2 border-dashed border-slate-200 hover:border-red-500 bg-white hover:bg-slate-50/50 rounded-xl py-4 px-4 text-center cursor-pointer transition-all duration-150 flex flex-col items-center justify-center group shadow-2xs"
                        >
                          <Upload className="w-6 h-6 text-slate-400 group-hover:text-red-500 mb-1.5 group-hover:scale-105 transition-transform" />
                          <span className="text-xs font-bold text-slate-700">
                            Upload Logo Image (PNG, JPG, SVG)
                          </span>
                          <span className="text-[10px] text-slate-400 mt-1 leading-normal">
                            Suggested aspect ratio approx. 7:4 (42mm &times; 24mm size constraints on printed media)
                          </span>
                        </div>
                        <input
                          type="file"
                          ref={logoInputRef}
                          onChange={handleLogoUpload}
                          accept="image/*"
                          className="hidden"
                        />
                      </div>
                    )}
                  </div>

                  <div>
                    <label className="block text-[10px] font-extrabold uppercase tracking-wider text-slate-500 mb-1">
                      Company Brand Name:
                    </label>
                    <input
                      type="text"
                      value={activeSender.name}
                      onChange={(e) => {
                        setSenders((prev) =>
                          prev.map((s) => (s.id === activeSenderId ? { ...s, name: e.target.value } : s))
                        );
                      }}
                      className="w-full text-xs px-3 py-2.5 bg-slate-50 border border-slate-200 rounded-lg focus:outline-none focus:border-slate-900 focus:ring-2 focus:ring-slate-900/5 font-bold"
                    />
                  </div>

                  <div>
                    <label className="block text-[10px] font-extrabold uppercase tracking-wider text-slate-500 mb-1">
                      Sender Location Detail Lines:
                    </label>
                    <textarea
                      rows={4}
                      value={activeSender.address}
                      onChange={(e) => {
                        setSenders((prev) =>
                          prev.map((s) => (s.id === activeSenderId ? { ...s, address: e.target.value } : s))
                        );
                      }}
                      className="w-full text-xs px-3 py-2.5 bg-slate-50 border border-slate-200 rounded-lg focus:outline-none focus:border-slate-900 focus:ring-2 focus:ring-slate-900/5 font-sans font-medium leading-relaxed"
                    />
                    <span className="text-[10px] text-slate-400 mt-1 block">
                      Use new lines to automatically format spacing lines. Only first 3 lines recommended.
                    </span>
                  </div>
                </div>

                {/* Configuration preferences space */}
                <div className="space-y-5">
                  <h4 className="text-[10px] font-extrabold text-slate-400 uppercase tracking-widest flex items-center gap-1.5">
                    <Sliders className="w-3.5 h-3.5 text-slate-400" />
                    <span>Styles &amp; Typography Calibration</span>
                  </h4>

                  {/* Selector: Custom Font family */}
                  <div className="bg-slate-50/50 p-4 border border-slate-200/60 rounded-xl space-y-2.5 shadow-2xs">
                    <label className="block text-[10px] font-extrabold uppercase tracking-wider text-slate-500">
                      Unified Envelope Font:
                    </label>
                    <select
                      value={envelopeFont}
                      onChange={(e) => setEnvelopeFont(e.target.value)}
                      className="w-full text-xs bg-white p-2.5 border border-slate-200 rounded-lg focus:outline-none focus:border-slate-900 font-bold text-slate-700"
                    >
                      <option value="Arial, Helvetica, sans-serif">Arial / Helvetica (Default Swiss)</option>
                      <option value='"Plus Jakarta Sans", sans-serif'>Plus Jakarta Sans</option>
                      <option value='"Inter", sans-serif'>Inter</option>
                      <option value='"Georgia", serif'>Georgia (Elegant Serif)</option>
                      <option value='"Courier New", monospace'>Courier New (Raw Polytype)</option>
                    </select>
                    <span className="text-[10px] text-slate-400 block leading-tight">
                      Ensures every text element inside the preview and physical print layout shares the exact same custom typeface.
                    </span>
                  </div>

                  {/* Printer Presets Configuration Block */}
                  <div className="bg-slate-50/50 p-4 border border-slate-200/60 rounded-xl space-y-4 shadow-2xs">
                    <div className="flex items-center justify-between">
                      <label className="block text-[10px] font-extrabold uppercase tracking-wider text-slate-500">
                        Printer Alignment Presets:
                      </label>
                      {selectedPresetId === "custom" && (
                        <span className="text-[9px] font-mono font-bold text-amber-600 bg-amber-50 px-1.5 py-0.5 rounded border border-amber-100 uppercase animate-pulse">
                          Unsaved Changes
                        </span>
                      )}
                    </div>

                    <div className="flex gap-2 items-center">
                      <select
                        value={selectedPresetId}
                        onChange={(e) => handleSelectPreset(e.target.value)}
                        className="flex-1 text-xs bg-white p-2.5 border border-slate-200 rounded-lg focus:outline-none focus:border-slate-900 font-bold shadow-2xs"
                      >
                        <option value="custom" disabled={selectedPresetId !== "custom"}>
                          {selectedPresetId === "custom" ? "Custom Adjustments (Sliders)" : "-- Select Preset --"}
                        </option>
                        {presets.map((p) => (
                          <option key={p.id} value={p.id}>
                            {p.name} {p.isBuiltIn ? "(Built-in)" : "(Saved)"}
                          </option>
                        ))}
                      </select>

                      {/* Delete / Rename Custom Preset Button */}
                      {selectedPresetId !== "custom" && !presets.find((p) => p.id === selectedPresetId)?.isBuiltIn && (
                        <div className="flex gap-1 shrink-0">
                          <button
                            onClick={() => handleRenamePreset(selectedPresetId)}
                            className="p-2.5 bg-white hover:bg-slate-50 text-slate-700 rounded-lg cursor-pointer transition-all border border-slate-205 shadow-2xs"
                            title="Rename selected preset"
                          >
                            <Edit2 className="w-3.5 h-3.5" />
                          </button>
                          <button
                            onClick={() => handleDeletePreset(selectedPresetId)}
                            className="p-2.5 bg-red-50 hover:bg-red-650 hover:text-white text-red-650 rounded-lg cursor-pointer transition-all border border-red-200 shadow-2xs"
                            title="Delete selected preset"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>
                        </div>
                      )}
                    </div>

                    {/* Quick Alignment Description for current preset */}
                    <div className="text-[10px] text-slate-400 bg-white p-2.5 rounded-lg border border-slate-200/50 space-y-1">
                      <div className="flex justify-between">
                        <span>Vertical Position:</span>
                        <span className="font-mono font-bold text-slate-600">{toBlockTopPadding}mm</span>
                      </div>
                      <div className="flex justify-between">
                        <span>Horizontal Shift:</span>
                        <span className="font-mono font-bold text-slate-600">{toBlockXShift}mm</span>
                      </div>
                      <div className="flex justify-between">
                        <span>Max Width:</span>
                        <span className="font-mono font-bold text-slate-600">{toBlockWidth}mm</span>
                      </div>
                    </div>

                    {/* Action to save current slider settings as new preset */}
                    <button
                      onClick={handleSaveCurrentAsPreset}
                      className="w-full text-xs text-red-600 hover:text-white font-extrabold bg-red-50 hover:bg-red-600 p-2.5 rounded-lg border border-red-200 hover:border-red-600 cursor-pointer flex items-center justify-center gap-1.5 transition-all shadow-2xs"
                    >
                      <Plus className="w-3.5 h-3.5" />
                      <span>Save Current Alignments as Preset</span>
                    </button>
                  </div>

                  {/* TO Recipient Address Block Calibration Widget */}
                  <div className="bg-slate-50/50 p-4 border border-slate-200/60 rounded-xl space-y-4 shadow-2xs">
                    <label className="block text-[10px] font-extrabold uppercase tracking-wider text-slate-500">
                      To Address Area Alignment &amp; Fine-Tuning:
                    </label>

                    {/* Slider: Top Padding (Vertical Offset) */}
                    <div>
                      <div className="flex justify-between items-center mb-1">
                        <span className="text-xs font-bold text-slate-700">
                          Vertical Position (Top Space):
                        </span>
                        <span className="text-xs font-mono font-bold text-red-600 bg-white border border-red-100 px-1.5 py-0.5 rounded">
                          {toBlockTopPadding}mm
                        </span>
                      </div>
                      <input
                        type="range"
                        min="5"
                        max="40"
                        step="1"
                        value={toBlockTopPadding}
                        onChange={(e) => setToBlockTopPadding(Number(e.target.value))}
                        className="w-full accent-red-600 h-1.5 bg-slate-200 rounded-lg cursor-pointer"
                      />
                    </div>

                    {/* Slider: Horizontal Shift */}
                    <div>
                      <div className="flex justify-between items-center mb-1">
                        <span className="text-xs font-bold text-slate-700">
                          Horizontal Shift (Offset X):
                        </span>
                        <span className="text-xs font-mono font-bold text-red-600 bg-white border border-red-100 px-1.5 py-0.5 rounded">
                          {toBlockXShift > 0 ? `+${toBlockXShift}` : toBlockXShift}mm
                        </span>
                      </div>
                      <input
                        type="range"
                        min="-30"
                        max="30"
                        step="1"
                        value={toBlockXShift}
                        onChange={(e) => setToBlockXShift(Number(e.target.value))}
                        className="w-full accent-red-600 h-1.5 bg-slate-200 rounded-lg cursor-pointer"
                      />
                    </div>

                    {/* Slider: Max Width */}
                    <div>
                      <div className="flex justify-between items-center mb-1">
                        <span className="text-xs font-bold text-slate-700">
                          Max Print block Width:
                        </span>
                        <span className="text-xs font-mono font-bold text-red-600 bg-white border border-red-100 px-1.5 py-0.5 rounded">
                          {toBlockWidth}mm
                        </span>
                      </div>
                      <input
                        type="range"
                        min="100"
                        max="160"
                        step="2"
                        value={toBlockWidth}
                        onChange={(e) => setToBlockWidth(Number(e.target.value))}
                        className="w-full accent-red-600 h-1.5 bg-slate-200 rounded-lg cursor-pointer"
                      />
                    </div>
                  </div>

                  {/* Slider: Recipient Spacing */}
                  <div>
                    <div className="flex justify-between items-center mb-1">
                      <label className="text-xs font-bold text-slate-700">
                        Recipient Address Letter Spacing:
                      </label>
                      <span className="text-xs font-mono font-bold text-red-600 bg-red-50 px-2 py-0.5 rounded">
                        {letterSpacing}px
                      </span>
                    </div>
                    <input
                      type="range"
                      min="0"
                      max="10"
                      step="1"
                      value={letterSpacing}
                      onChange={(e) => setLetterSpacing(Number(e.target.value))}
                      className="w-full accent-red-600 h-1.5 bg-slate-100 rounded-lg cursor-pointer"
                    />
                    <span className="text-[10px] text-slate-400 mt-1 block leading-normal font-sans">
                      High letter-spacing matches the luxury/international style demonstrated in target mockup printout.
                    </span>
                  </div>

                  {/* Slider: Sender Spacing */}
                  <div>
                    <div className="flex justify-between items-center mb-1">
                      <label className="text-xs font-bold text-slate-700">
                        Sender Letter Spacing:
                      </label>
                      <span className="text-xs font-mono font-bold text-red-600 bg-red-50 px-2 py-0.5 rounded">
                        {fromLetterSpacing}px
                      </span>
                    </div>
                    <input
                      type="range"
                      min="0"
                      max="8"
                      step="1"
                      value={fromLetterSpacing}
                      onChange={(e) => setFromLetterSpacing(Number(e.target.value))}
                      className="w-full accent-red-600 h-1.5 bg-slate-100 rounded-lg cursor-pointer"
                    />
                  </div>

                  {/* Toggle Option: To prefix */}
                  <div className="flex items-center justify-between p-3.5 bg-slate-50 border border-slate-200/60 rounded-xl shadow-2xs">
                    <div>
                      <span className="text-xs font-bold text-slate-800 block">Show "To:" prefix</span>
                      <span className="text-[10px] text-slate-400 font-medium">Toggle "To:" header on recipients block</span>
                    </div>
                    <button
                      onClick={() => setShowToPrefix(!showToPrefix)}
                      className={`w-11 h-6 rounded-full p-0.5 transition-colors cursor-pointer ${
                        showToPrefix ? "bg-red-650" : "bg-slate-200"
                      }`}
                    >
                      <div
                        className={`bg-white w-5 h-5 rounded-full shadow-sm transform transition-transform ${
                          showToPrefix ? "translate-x-5" : "translate-x-0"
                        }`}
                      />
                    </button>
                  </div>

                  {/* Stamp text selection */}
                  <div>
                    <label className="text-xs font-bold text-slate-700 block mb-1">
                      Stamp Box Value:
                    </label>
                    <input
                      type="text"
                      className="w-full text-xs px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-none"
                      value={stampText}
                      onChange={(e) => setStampText(e.target.value)}
                      placeholder="e.g. POSTAGE PAID, BOOK-POST"
                    />
                    <div className="flex items-center gap-1.5 mt-2">
                      <input
                        type="checkbox"
                        id="stampborder"
                        checked={enableStampBorder}
                        onChange={(e) => setEnableStampBorder(e.target.checked)}
                        className="rounded text-red-600 accent-red-600"
                      />
                      <label htmlFor="stampborder" className="text-[11px] text-slate-500 font-bold">
                        Render outer stamp boundary box
                      </label>
                    </div>
                  </div>

                </div>

              </div>
            </div>
          )}

        </main>
      </div>

      {/* Hidden real page blocks specifically for physical print engines */}
      <div className="print-only-wrapper bg-white text-black min-h-screen">
        {schools
          .filter((school) => school.selected)
          .map((school) => (
            <div
              key={school.id}
              className="print-container"
              style={{
                width: "265mm",
                height: "112mm",
                padding: "5mm 8mm 6mm 4.5mm",
                boxSizing: "border-box",
                overflow: "hidden",
                fontFamily: envelopeFont
              }}
            >
              {/* Top Row: Sender block and Stamp box */}
              <div className="flex justify-between items-start" style={{ height: "27mm" }}>
                 <div className="flex items-start">
                  {activeSender.logoUrl ? (
                    <div
                      className="mr-3 shrink-0 border border-slate-100 bg-white rounded-md p-1 flex items-center justify-center overflow-hidden"
                      style={{
                        width: "42mm",
                        height: "24mm",
                        minWidth: "42mm",
                        minHeight: "24mm"
                      }}
                    >
                      <img
                        src={activeSender.logoUrl}
                        alt="Brand Logo"
                        style={{
                          maxWidth: "100%",
                          maxHeight: "100%",
                          width: "auto",
                          height: "auto",
                          objectFit: "contain"
                        }}
                        referrerPolicy="no-referrer"
                      />
                    </div>
                  ) : (
                    /* B2P styled Logo representation */
                    <svg
                      viewBox="0 0 100 40"
                      className="mr-3 shrink-0 block"
                      style={{
                        width: "42mm",
                        height: "24mm",
                        minWidth: "42mm",
                        minHeight: "24mm",
                        fontFamily: envelopeFont
                      }}
                    >
                      <path
                        d="M12 4 a8 8 0 0 1 8 8 v14 a8 8 0 0 1 -16 0 v-20 h5 v6 a8 8 0 0 1 3 0 a8 8 0 0 1 0 0 v8"
                        fill="none"
                        stroke="#DC2626"
                        strokeWidth="4.5"
                        strokeLinecap="round"
                      />
                      <text x="36" y="24" fill="#000000" fontStyle="oblique" fontSize="17" fontWeight="900" letterSpacing="1">B2P</text>
                      <text x="36" y="34" fill="#4B5563" fontSize="8" fontWeight="600" letterSpacing="0.5">International</text>
                    </svg>
                  )}

                  {/* Core Sender Info address block */}
                  <div className="from-details text-left shrink-0" style={{ marginLeft: "7mm", paddingTop: "1.8mm" }}>
                    <span
                      className="from-label block text-[#090909] font-extrabold uppercase"
                      style={{
                        fontSize: "12pt",
                        lineHeight: "1",
                        letterSpacing: `${fromLetterSpacing}px`,
                        marginBottom: "3mm",
                        fontFamily: envelopeFont
                      }}
                    >
                      FROM:
                    </span>
                    
                    <div
                      className="from-name text-[#151515] font-extrabold"
                      style={{
                        fontSize: "11.5pt",
                        lineHeight: "1.1",
                        letterSpacing: `${fromLetterSpacing}px`,
                        marginBottom: "2.2mm",
                        fontFamily: envelopeFont
                      }}
                    >
                      {activeSender.name}
                    </div>

                    <div
                      className="from-addr text-[#4a4a4a] font-normal"
                      style={{
                        fontSize: "9.5pt",
                        lineHeight: "1.55",
                        letterSpacing: "0px",
                        fontFamily: envelopeFont
                      }}
                    >
                      {activeSender.address.split("\n").map((line, idx) => (
                        <span key={idx} className="block whitespace-nowrap">
                          {line}
                        </span>
                      ))}
                    </div>
                  </div>
                </div>

                {/* Postage Stamp Box coordinates match */}
                {enableStampBorder ? (
                  <div
                    className="border border-black flex items-center justify-center shrink-0"
                    style={{
                      width: "29.5mm",
                      height: "25mm",
                      marginTop: "1.5mm"
                    }}
                  >
                    <span
                      className="text-black text-center uppercase whitespace-pre-wrap leading-relaxed tracking-wider"
                      style={{
                        fontSize: "6.5pt",
                        fontWeight: 700,
                        letterSpacing: "1.5px",
                        fontFamily: envelopeFont
                      }}
                    >
                      {stampText}
                    </span>
                  </div>
                ) : (
                  <div
                    className="flex items-center justify-center shrink-0 text-right"
                    style={{
                      width: "29.5mm",
                      height: "25mm",
                      marginTop: "1.5mm"
                    }}
                  >
                    <span
                      className="text-black text-center uppercase whitespace-pre-wrap leading-relaxed tracking-wider"
                      style={{
                        fontSize: "6.5pt",
                        fontWeight: 700,
                        letterSpacing: "1.5px",
                        fontFamily: envelopeFont
                      }}
                    >
                      {stampText}
                    </span>
                  </div>
                )}
              </div>

              {/* Center spacer block matching exact visual layout request */}
              <div style={{ height: "4.7mm" }} />

              {/* To Recipient location matching requested design */}
              <div
                className="to-block flex flex-col items-start text-left"
                style={{
                  width: `${toBlockWidth}mm`,
                  marginLeft: "auto",
                  paddingTop: `${toBlockTopPadding}mm`,
                  transform: `translateX(${toBlockXShift}mm)`
                }}
              >
                {showToPrefix && (
                  <span
                    className="to-label block font-extrabold text-[#050505] text-left"
                    style={{
                      fontSize: "14pt",
                      lineHeight: "1",
                      letterSpacing: `${letterSpacing}px`,
                      marginBottom: "6mm",
                      fontFamily: envelopeFont
                    }}
                  >
                    To:
                  </span>
                )}

                <div
                  className="to-name text-black font-extrabold text-left capitalize tracking-normal"
                  style={{
                    fontSize: "13.5pt",
                    lineHeight: "1.45",
                    letterSpacing: `${letterSpacing}px`,
                    marginBottom: "1.5mm",
                    whiteSpace: "normal",
                    fontFamily: envelopeFont
                  }}
                >
                  {school.name}
                </div>

                <div
                  className="to-addr text-black font-medium text-left tracking-normal whitespace-pre-line leading-relaxed"
                  style={{
                    fontSize: "13pt",
                    lineHeight: "1.45",
                    letterSpacing: `${letterSpacing}px`,
                    fontFamily: envelopeFont
                  }}
                >
                  {school.address}
                </div>
              </div>
            </div>
          ))}
      </div>

      {/* Primary Edit School Modal Dialog box Backdrop */}
      {editingSchool && (
        <div className="fixed inset-0 bg-slate-900/60 z-50 flex items-center justify-center p-4 backdrop-blur-sm transition-all duration-200">
          <div className="bg-white rounded-xl shadow-xl border border-slate-200 max-w-md w-full p-6 animate-in fade-in zoom-in duration-150">
            
            <div className="flex items-center justify-between border-b border-slate-100 pb-3.5 mb-4">
              <span className="text-sm font-extrabold text-slate-800 flex items-center gap-1.5">
                <Edit2 className="w-4.5 h-4.5 text-red-600" />
                <span>Modify Address Row</span>
              </span>
              <button
                onClick={() => setEditingSchool(null)}
                className="text-slate-400 hover:text-slate-600 cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="space-y-4">
              <div>
                <label className="block text-[10px] font-extrabold text-slate-500 uppercase tracking-widest mb-1">
                  School / Recipient Name:
                </label>
                <input
                  type="text"
                  value={editingSchool.name || ""}
                  onChange={(e) =>
                    setEditingSchool({ ...editingSchool, name: e.target.value })
                  }
                  className="w-full text-xs px-3.5 py-2.5 border border-slate-200 rounded-lg focus:outline-none focus:border-red-500"
                />
              </div>

              <div>
                <div className="flex justify-between items-center mb-1">
                  <label className="block text-[10px] font-extrabold text-slate-500 uppercase tracking-widest">
                    Recipient Address coordinates:
                  </label>
                  <button
                    type="button"
                    onClick={() => {
                      setEditingSchool({
                        ...editingSchool,
                        address: autoFormatAddressText(editingSchool.address || "")
                      });
                    }}
                    className="text-[9px] font-extrabold text-red-600 hover:text-red-700 hover:bg-red-50 px-2 py-0.5 rounded border border-red-100 flex items-center gap-1 cursor-pointer transition-all uppercase"
                    title="Convert to Title Case and clean zip code formatting"
                  >
                    <Sparkles className="w-2.5 h-2.5" />
                    <span>Auto-Format</span>
                  </button>
                </div>
                <textarea
                  rows={4}
                  value={editingSchool.address || ""}
                  onChange={(e) =>
                    setEditingSchool({ ...editingSchool, address: e.target.value })
                  }
                  className="w-full text-xs px-3.5 py-2.5 border border-slate-200 rounded-lg focus:outline-none focus:border-red-500 font-sans resize-none"
                />
              </div>
            </div>

            <div className="mt-5 pt-3 border-t border-slate-100 flex justify-end gap-2 text-xs">
              <button
                onClick={() => setEditingSchool(null)}
                className="px-4 py-2.5 bg-slate-50 hover:bg-slate-100 border border-slate-200 text-slate-600 rounded-lg font-bold cursor-pointer transition-all"
              >
                Cancel
              </button>
              <button
                onClick={handleSaveEdit}
                disabled={!editingSchool.name?.trim() || !editingSchool.address?.trim()}
                className="px-5 py-2.5 bg-red-600 hover:bg-red-700 text-white rounded-lg font-extrabold shadow-sm transition-all cursor-pointer"
              >
                Save Changes &amp; Close
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Custom dialog system to bypass blocked dialog alerts/confirms/prompts in iframe */}
      {dialog.isOpen && (
        <div className="fixed inset-0 bg-slate-950/60 backdrop-blur-xs flex items-center justify-center p-4 z-[9999] no-print animate-fade-in animate-duration-150">
          <div className="bg-white rounded-2xl shadow-2xl border border-slate-100 max-w-md w-full p-6 space-y-4 relative overflow-hidden select-none">
            
            {/* Top background color accent depending on action type */}
            <div className={`absolute top-0 left-0 right-0 h-1.5 ${
              dialog.type === 'confirm' ? 'bg-red-500' :
              dialog.type === 'alert' ? 'bg-amber-500' : 'bg-indigo-500'
            }`} />

            {/* Header info */}
            <div className="flex items-start gap-3">
              <div className={`p-2 rounded-xl shrink-0 ${
                dialog.type === 'confirm' ? 'bg-red-50 text-red-600' :
                dialog.type === 'alert' ? 'bg-amber-50 text-amber-600' : 'bg-indigo-50 text-indigo-600'
              }`}>
                {dialog.type === 'confirm' ? (
                  <Trash2 className="w-5 h-5" />
                ) : dialog.type === 'alert' ? (
                  <Info className="w-5 h-5" />
                ) : (
                  <Edit2 className="w-5 h-5" />
                )}
              </div>
              <div className="space-y-1">
                <h3 className="text-sm font-black text-slate-900 tracking-tight">
                  {dialog.title}
                </h3>
                <p className="text-xs text-slate-500 leading-relaxed font-sans">
                  {dialog.message}
                </p>
              </div>
            </div>

            {/* Inputs based on types */}
            {(dialog.type === 'prompt' || dialog.type === 'double_prompt') && (
              <div className="space-y-3.5 pt-1">
                <div>
                  <label className="block text-[9px] font-extrabold text-slate-400 uppercase tracking-widest mb-1 font-mono">
                    {dialog.inputLabel || "Input Value"}
                  </label>
                  <input
                    type="text"
                    value={dialogInput1}
                    onChange={(e) => setDialogInput1(e.target.value)}
                    className="w-full text-xs px-3.5 py-2.5 border border-slate-200 rounded-lg focus:outline-none focus:border-red-500 font-sans shadow-3xs"
                    autoFocus
                  />
                </div>

                {dialog.type === 'double_prompt' && (
                  <div>
                    <label className="block text-[9px] font-extrabold text-slate-400 uppercase tracking-widest mb-1 font-mono">
                      {dialog.inputLabelSecond || "Extended Details"}
                    </label>
                    <textarea
                      rows={4}
                      value={dialogInput2}
                      onChange={(e) => setDialogInput2(e.target.value)}
                      className="w-full text-xs px-3.5 py-2.5 border border-slate-200 rounded-lg focus:outline-none focus:border-red-500 font-sans resize-none shadow-3xs"
                    />
                  </div>
                )}
              </div>
            )}

            {/* Action buttons */}
            <div className="flex items-center justify-end gap-2 text-xs pt-2">
              {dialog.type !== 'alert' && (
                <button
                  onClick={() => setDialog((prev) => ({ ...prev, isOpen: false }))}
                  className="px-4 py-2 bg-slate-50 hover:bg-slate-100 border border-slate-200 text-slate-650 rounded-lg font-bold cursor-pointer transition-all hover:text-slate-900"
                >
                  Cancel
                </button>
              )}
              <button
                onClick={() => {
                  dialog.onConfirm(dialogInput1, dialogInput2);
                  setDialog((prev) => ({ ...prev, isOpen: false }));
                }}
                className={`px-4.5 py-2 rounded-lg font-extrabold shadow-3xs transition-all cursor-pointer text-white ${
                  dialog.type === 'confirm' ? 'bg-red-600 hover:bg-red-700' :
                  dialog.type === 'alert' ? 'bg-amber-600 hover:bg-amber-700' : 'bg-red-600 hover:bg-red-700'
                }`}
              >
                {dialog.type === 'confirm' ? 'Confirm Action' : 'Proceed'}
              </button>
            </div>

          </div>
        </div>
      )}
    </div>
  );
}
