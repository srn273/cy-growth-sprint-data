import { useState, useEffect, useRef } from "react";

export default function SprintDashboard() {
  const [isEditMode, setIsEditMode] = useState(false);
  const [isPresentMode, setIsPresentMode] = useState(false);
  const [currentSlide, setCurrentSlide] = useState(1);
  const [currentSprint, setCurrentSprint] = useState(0);
  const [showImportModal, setShowImportModal] = useState(false);
  const [currentImportSlideId, setCurrentImportSlideId] = useState<number | null>(null);
  const [pastedData, setPastedData] = useState("");
  const [pastedImage, setPastedImage] = useState<string | null>(null);
  const [isProcessingImage, setIsProcessingImage] = useState(false);
  const [editingTableId, setEditingTableId] = useState<number | null>(null);
  const [editingStatsId, setEditingStatsId] = useState<number | null>(null);
  const [isGlobalImport, setIsGlobalImport] = useState(false);

  // Track scroll position to prevent jumping
  const scrollPositionRef = useRef<number>(0);
  const isPreservingScrollRef = useRef<boolean>(false);

  const MAX_SPRINTS = 5;

  // Preserve scroll position during state updates to prevent jumping
  useEffect(() => {
    if (isPreservingScrollRef.current) {
      window.scrollTo(0, scrollPositionRef.current);
      isPreservingScrollRef.current = false;
    }
  });

  // Save scroll position before state updates
  const preserveScroll = () => {
    scrollPositionRef.current = window.scrollY;
    isPreservingScrollRef.current = true;
  };

  // Helper function to recalculate totals and stats for a slide after data changes
  const recalculateSlideStats = (slide: any) => {
    if (!slide || !slide.data) return;

    const recalculateTableStats = (data: any) => {
      if (!data.rows || data.rows.length === 0) return;

      // Recalculate totals for tables with total row
      if (data.total && data.columns) {
        const newTotal: any = {};
        data.columns.forEach((col: any) => {
          if (col.key !== "sprint" && col.key !== "metrics" && col.key !== "quarter") {
            const sum = data.rows.reduce((acc: number, row: any) => {
              const val = parseFloat(row[col.key]) || 0;
              return acc + val;
            }, 0);
            newTotal[col.key] = sum;
          }
        });
        Object.assign(data.total, newTotal);
      }

      // Recalculate quarter stats for quarterStats type
      if (data.quarterStats && data.columns) {
        const newStats: any = {};
        data.columns.forEach((col: any) => {
          if (col.key !== "sprint") {
            const sum = data.rows.reduce((acc: number, row: any) => {
              const val = parseFloat(row[col.key]) || 0;
              return acc + val;
            }, 0);
            newStats[col.key] = sum;
          }
        });
        Object.assign(data.quarterStats, newStats);
      }
    };

    // Handle different slide types
    switch (slide.type) {
      case "supportData":
        recalculateTableStats(slide.data.tickets);
        recalculateTableStats(slide.data.liveChat);
        break;
      case "agencyLeads":
        recalculateTableStats(slide.data.leadsConversion);
        recalculateTableStats(slide.data.q3Performance);
        break;
      case "wixApp":
        // Recalculate lifetime stats for Wix App
        if (slide.data.lifetime && slide.data.rows && slide.data.columns) {
          const newLifetime: any = {};
          slide.data.columns.forEach((col: any) => {
            if (col.key !== "sprint" && col.key in slide.data.lifetime) {
              const sum = slide.data.rows.reduce((acc: number, row: any) => {
                const val = parseFloat(row[col.key]) || 0;
                return acc + val;
              }, 0);
              newLifetime[col.key] = sum;
            }
          });
          Object.assign(slide.data.lifetime, newLifetime);
        }
        recalculateTableStats(slide.data);
        break;
      case "referral":
        // Recalculate lifetime stats for Referral
        if (slide.data.lifetime && slide.data.rows && slide.data.columns) {
          const newLifetime: any = {};
          slide.data.columns.forEach((col: any) => {
            if (col.key !== "sprint" && col.key in slide.data.lifetime) {
              const sum = slide.data.rows.reduce((acc: number, row: any) => {
                const val = parseFloat(row[col.key]) || 0;
                return acc + val;
              }, 0);
              newLifetime[col.key] = sum;
            }
          });
          Object.assign(slide.data.lifetime, newLifetime);
        }
        recalculateTableStats(slide.data);
        break;
      case "withTarget":
      case "quarterStats":
      default:
        recalculateTableStats(slide.data);
        break;
    }
  };

  const maintainSprintLimit = (rows) => {
    if (!rows || rows.length <= MAX_SPRINTS) return rows;

    const isSpecial = (r: any) => {
      const s = String(r?.sprint ?? "").toLowerCase();
      return s === "qtd" || s === "total" || s === "lifetime" || s.includes("qtd");
    };

    const specialRows = rows.filter(isSpecial);
    const numericRows = rows.filter((r) => !isSpecial(r));

    // Sort by sprint number (descending) and keep only last 5 numeric sprints
    const sorted = [...numericRows].sort((a, b) => {
      const sprintA = typeof a.sprint === "number" ? a.sprint : parseInt(a.sprint) || 0;
      const sprintB = typeof b.sprint === "number" ? b.sprint : parseInt(b.sprint) || 0;
      return sprintB - sprintA;
    });

    return [...sorted.slice(0, MAX_SPRINTS), ...specialRows];
  };

  // Helper: check if a slide type uses nested tables
  const isNestedTableSlide = (type: string): boolean => {
    return ["supportData", "agencyLeads", "referral", "wixApp", "subscriptions"].includes(type);
  };

  // Helper: resolve parent slide id from nested table ids (e.g., +1000 for support, +2000 for agency)
  const getActualSlideId = (slideId: number) => {
    if (slideId >= 2000) return slideId - 2000;
    if (slideId >= 1000) return slideId - 1000;
    return slideId;
  };

  const parseStatsData = (lines: string[], slideId: number) => {
    // Detect which slide type and get expected stat keys
    const slide = sprintData.slides.find((s) => s.id === slideId);
    if (!slide) return null;

    let expectedKeys: string[] = [];
    if (slide.type === "quarterStats" && slide.data.quarterStats) {
      expectedKeys = Object.keys(slide.data.quarterStats);
    } else if (slide.type === "withTarget" && slide.data.total) {
      expectedKeys = Object.keys(slide.data.total);
    } else if (slide.type === "referral" && slide.data.lifetime) {
      expectedKeys = Object.keys(slide.data.lifetime);
    } else if (slide.type === "wixApp" && slide.data.lifetime) {
      expectedKeys = Object.keys(slide.data.lifetime);
    }

    if (expectedKeys.length === 0) return null;

    const normalize = (s: string) =>
      String(s)
        .toLowerCase()
        .replace(/[^a-z0-9]/g, "");
    const statsData: Record<string, number> = {};

    // Try to detect format
    const firstLine = lines[0];

    // Option 1 & 2: Key-value format (tab/colon separated or two-column)
    if (firstLine.includes("\t") || firstLine.includes(":")) {
      const delimiter = firstLine.includes("\t") ? "\t" : ":";
      lines.forEach((line) => {
        const parts = line.split(delimiter).map((p) => p.trim());
        if (parts.length >= 2) {
          const key = normalize(parts[0]);
          const value = parseFloat(parts[1].replace(/[$,%]/g, "").replace(/,/g, ""));

          // Match to expected keys
          const matchedKey = expectedKeys.find(
            (k) => normalize(k) === key || key.includes(normalize(k)) || normalize(k).includes(key),
          );
          if (matchedKey && !isNaN(value)) {
            statsData[matchedKey] = value;
          }
        }
      });
    }
    // Option 3: Just values (one per line, in order)
    else if (lines.every((line) => !isNaN(parseFloat(line.replace(/[$,%]/g, "").replace(/,/g, ""))))) {
      lines.forEach((line, idx) => {
        if (idx < expectedKeys.length) {
          const value = parseFloat(line.replace(/[$,%]/g, "").replace(/,/g, ""));
          if (!isNaN(value)) {
            statsData[expectedKeys[idx]] = value;
          }
        }
      });
    }

    return Object.keys(statsData).length > 0 ? statsData : null;
  };

  const handleImageExtraction = async (imageData: string) => {
    setIsProcessingImage(true);
    try {
      const response = await fetch(`${import.meta.env.VITE_SUPABASE_URL}/functions/v1/extract-screenshot-data`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${import.meta.env.VITE_SUPABASE_PUBLISHABLE_KEY}`,
        },
        body: JSON.stringify({ imageData }),
      });

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error || "Failed to extract data from screenshot");
      }

      const result = await response.json();
      const extractedText = result.extractedText || "";

      if (extractedText.trim()) {
        setPastedData(extractedText);
        setPastedImage(null);
        alert("✅ Data extracted from screenshot! Review and click Import Data to proceed.");
      } else {
        alert("Could not extract data from the image. Please try pasting text data directly.");
      }
    } catch (error) {
      console.error("Error processing image:", error);
      alert(
        error instanceof Error ? error.message : "Error processing screenshot. Please try pasting text data directly.",
      );
    } finally {
      setIsProcessingImage(false);
    }
  };

  const handlePasteEvent = (e: React.ClipboardEvent) => {
    const items = e.clipboardData?.items;
    if (!items) return;

    for (let i = 0; i < items.length; i++) {
      if (items[i].type.indexOf("image") !== -1) {
        e.preventDefault();
        const blob = items[i].getAsFile();
        if (blob) {
          const reader = new FileReader();
          reader.onload = (event) => {
            const imageData = event.target?.result as string;
            setPastedImage(imageData);
            handleImageExtraction(imageData);
          };
          reader.readAsDataURL(blob);
        }
        return;
      }
    }
  };

  const handlePastedData = () => {
    if (!pastedData.trim()) {
      alert("Please paste data from your Google Sheet or upload a screenshot first.");
      return;
    }

    try {
      console.log("Raw pasted data:", pastedData);
      const lines = pastedData.split("\n").filter((row) => row.trim());

      if (lines.length < 1) {
        alert("Please paste at least one row of data.");
        return;
      }

      const firstLine = lines[0];

      // Detect if it's table data or stats data
      // Table data: Has tabs (multiple columns) AND multiple rows with consistent column count
      // Stats data: Single column OR key-value pairs
      const hasMultipleColumns = firstLine.includes("\t") && firstLine.split("\t").length > 2;
      const isMultipleRows = lines.length >= 2;

      // If it looks like table data (multiple columns and rows), parse as table
      if (hasMultipleColumns && isMultipleRows) {
        console.log("Detected as TABLE data");

        const delimiter = "\t"; // Google Sheets uses tabs
        console.log("Using TAB delimiter for Google Sheets data");

        const headers = lines[0].split(delimiter).map((h) => h.trim().toLowerCase());
        console.log("Parsed headers:", headers);

        if (headers.length < 2) {
          alert("Table data must have at least 2 columns. Please check your data.");
          return;
        }

        const dataRows = lines.slice(1).map((line) => {
          const values = line.split(delimiter).map((v) => v.trim());
          const row: any = {};
          headers.forEach((header, idx) => {
            const value = values[idx] || "";
            // Clean and convert numbers (remove $, %, commas)
            const cleaned = String(value).replace(/[$,%]/g, "").replace(/,/g, "");
            const numValue = Number(cleaned);
            row[header] = !isNaN(numValue) && cleaned !== "" ? numValue : value;
          });
          return row;
        });
        console.log("Parsed data rows:", dataRows);

        // Apply 5-sprint limit
        const limitedRows = maintainSprintLimit(dataRows);
        console.log("Limited rows (max 5):", limitedRows);

        // Update the specific slide with imported data
        if (currentImportSlideId) {
          console.log("Updating slide ID:", currentImportSlideId);
          importDataToSlide(currentImportSlideId, headers, limitedRows);
        }

        alert(`Table data imported successfully! ${limitedRows.length} row(s) loaded.`);
        setShowImportModal(false);
        setCurrentImportSlideId(null);
        setPastedData("");
        return;
      }

      // Otherwise, try to parse as stats data
      if (currentImportSlideId) {
        console.log("Detected as STATS data");
        const statsData = parseStatsData(lines, currentImportSlideId);

        if (statsData) {
          console.log("Parsed as stats data:", statsData);
          importStatsToSlide(currentImportSlideId, statsData);
          alert(`Stats imported successfully! ${Object.keys(statsData).length} stat(s) updated.`);
          setShowImportModal(false);
          setCurrentImportSlideId(null);
          setPastedData("");
          return;
        }
      }

      alert(
        "Unable to parse data. Please ensure:\n\n• For tables: Copy the entire table from Google Sheets including headers\n• For stats: Use one of the 3 supported formats\n\nSee instructions above for details.",
      );
    } catch (error) {
      console.error("Import error:", error);
      alert("Error parsing data. Please check the format and try again.");
    }
  };

  const importStatsToSlide = (slideId: number, statsData: Record<string, number>) => {
    setSprintData((prev) => ({
      ...prev,
      slides: prev.slides.map((slide) => {
        if (slide.id !== slideId) return slide;

        const newSlide = JSON.parse(JSON.stringify(slide));

        if (slide.type === "quarterStats" && newSlide.data.quarterStats) {
          Object.keys(statsData).forEach((key) => {
            if (newSlide.data.quarterStats[key] !== undefined) {
              newSlide.data.quarterStats[key] = statsData[key];
            }
          });
        } else if (slide.type === "withTarget" && newSlide.data.total) {
          Object.keys(statsData).forEach((key) => {
            if (newSlide.data.total[key] !== undefined) {
              newSlide.data.total[key] = statsData[key];
            }
          });
        } else if ((slide.type === "referral" || slide.type === "wixApp") && newSlide.data.lifetime) {
          Object.keys(statsData).forEach((key) => {
            if (newSlide.data.lifetime[key] !== undefined) {
              newSlide.data.lifetime[key] = statsData[key];
            }
          });
        }

        return newSlide;
      }),
    }));
  };

  const importDataToSlide = (slideId, headers, rows) => {
    const normalize = (s: string) =>
      String(s)
        .toLowerCase()
        .replace(/[^a-z0-9]/g, "");

    // Build a lookup from normalized header -> original header
    const headerLookup: Record<string, string> = {};
    headers.forEach((h: string) => {
      headerLookup[normalize(h)] = h;
    });

    const synonymMap: Record<string, string[]> = {
      // Common
      sprint: ["sprint", "sprintid", "sprintno", "sprintnumber", "sprint#"],
      // Content/brand
      blog: ["blog", "blogs", "blogposts", "blogpost", "blogarticles"],
      infographics: ["infographics", "infographic"],
      kb: ["kb", "kb articles", "kbarticles", "knowledgebase", "knowledge base", "docs", "documentation", "articles"],
      videos: ["videos", "video", "youtube", "yt"],
      total: ["total", "totalpaid", "paidtotal", "totalrevenue", "total paid"],
      direct: ["direct", "directplans", "direct plans"],
      social: ["social", "socialmedia"],
      blogs: ["blogs", "blogmentions", "blog mentions"],
      youtube: ["youtube", "yt"],
      negative: ["negative", "neg"],
      position: ["position", "pluginposition", "plugin position", "rank", "ranking"],

      // Plugin ranking & rankings
      pos1_2: ["position 1-2", "pos1_2", "pos12", "position1-2", "position12", "1-2", "12"],
      pos3_10: ["position 3-10", "pos3_10", "pos310", "position3-10", "position310", "3-10", "310"],

      // Support - tickets
      totalTickets: [
        "total tickets",
        "total tickets solved",
        "tickets",
        "tickets solved",
        "totaltickets",
        "totalticketsolved",
      ],
      avgFirstResponse: [
        "avg first response",
        "average first response",
        "first response",
        "avgfirstresponse",
        "firstresponse",
      ],
      csat: ["csat", "csat score", "customer satisfaction", "satisfaction"],
      presales: ["pre-sales tickets", "presales tickets", "pre sales tickets", "presales"],
      converted: ["converted tickets", "converted", "converted (unique customers)", "convertedunique"],
      paidSubs: ["total paid subscriptions (websites)", "paid subscriptions", "paid subs", "paidsubs"],
      agencyTickets: ["agency tickets", "agencysupport", "agency"],
      badRating: ["bad rating", "negative rating", "bad", "rating bad"],

      // Support - live chat
      conversations: ["conversations", "conversations assigned", "assigned conversations"],
      avgAssignment: [
        "avg teammate assignment to first response",
        "avg teammate assignment",
        "assignment to first response",
        "avgassignment",
      ],
      avgResolution: ["avg full resolution time", "avg resolution", "resolution time", "avgresolution"],

      // Agency leads
      metrics: ["metrics", "category", "metric"],
      totalCount: ["total count", "count", "totalcount", "total"],
      fromTickets: ["from tickets", "tickets", "fromtickets"],
      websiteLeads: ["website leads (lp)", "website leads", "website", "websiteleads"],
      fromAds: ["from ads", "ads", "paid ads", "fromads"],
      liveChat: ["live chat", "livechat", "chat"],
      webApp: ["web app (book a call)", "web app", "book a call", "webapp"],

      // Agency performance
      quarter: ["quarter", "sprint", "period"],
      target: ["target", "goal"],
      achieved: ["achieved", "actual"],
      percentage: ["percentage", "%", "percent", "targetachieved%", "target achieved %"],

      // withTarget types (Agency Signups, Affiliate)
      signups: ["signups", "signup", "new signups"],
      paid: ["paid", "paidusers", "paid users", "paying"],
      shortfall: ["shortfall", "shortage", "deficit"],

      // Affiliate
      trials: ["trial signups", "trialsignups", "trials"],
      conversions: ["conversions", "paid signups", "paidsignups"],

      // Referral
      referrals: ["referrals", "referrals generated", "referralsgenerated"],

      // Wix App
      freeSignups: ["free signups", "freesignups", "free"],
    };

    const findHeaderForKey = (key: string): string | null => {
      const keyN = normalize(key);
      // 1) exact key match
      if (headerLookup[keyN]) return headerLookup[keyN];
      // 2) synonyms
      const syns = synonymMap[key] || [];
      for (const s of syns) {
        const sN = normalize(s);
        if (headerLookup[sN]) return headerLookup[sN];
      }
      // 3) fuzzy contains
      const found = Object.keys(headerLookup).find((hN) => hN.includes(keyN) || keyN.includes(hN));
      return found ? headerLookup[found] : null;
    };

    const coerceNumber = (val: any) => {
      if (val === null || val === undefined) return val;
      const str = String(val).trim();
      // e.g., "$1,234", "45%", "1,234"
      const cleaned = str.replace(/[$,%]/g, "").replace(/,/g, "");
      const num = Number(cleaned);
      return isNaN(num) ? val : num;
    };

    setSprintData((prev) => ({
      ...prev,
      slides: prev.slides.map((slide) => {
        if (slide.id !== slideId) return slide;

        const newSlide = JSON.parse(JSON.stringify(slide));

        // Handle different slide types
        if (slide.type === "table" || slide.type === "pluginRanking") {
          if (newSlide.data?.columns && newSlide.data?.rows) {
            const expectedKeys: string[] = newSlide.data.columns.map((c: any) => c.key);

            const mapped = rows.map((r: any) => {
              const out: Record<string, any> = {};
              expectedKeys.forEach((k) => {
                const header = findHeaderForKey(k);
                let v = header ? r[header] : undefined;
                if (k === "sprint" && v !== undefined) {
                  // Allow values like "Sprint 263"
                  const parsed = parseInt(String(v).match(/\d+/)?.[0] || String(v), 10);
                  v = isNaN(parsed) ? v : parsed;
                } else {
                  v = coerceNumber(v);
                }
                out[k] = v !== undefined ? v : r[k];
              });
              return out;
            });

            newSlide.data.rows = mapped;
          }
        } else if (slide.type === "supportData") {
          // Auto-detect which table to update based on headers
          const normalizedHeaders = headers.map((h) => h.toLowerCase().replace(/[^a-z0-9]/g, ""));
          const hasTicketsHeaders = normalizedHeaders.some(
            (h) =>
              h.includes("totaltickets") ||
              h.includes("totalticketsolved") ||
              h.includes("presales") ||
              h.includes("converted") ||
              h.includes("paidsubs"),
          );
          const hasLiveChatHeaders = normalizedHeaders.some(
            (h) =>
              h.includes("conversations") ||
              h.includes("conversationsassigned") ||
              h.includes("avgassignment") ||
              h.includes("avgteammate"),
          );

          const targetKey = hasTicketsHeaders ? "tickets" : hasLiveChatHeaders ? "liveChat" : "tickets";
          const target = newSlide.data[targetKey];

          if (target?.columns) {
            const expectedKeys: string[] = target.columns.map((c: any) => c.key);
            const mapped = rows.map((r: any) => {
              const out: Record<string, any> = {};
              expectedKeys.forEach((k) => {
                const header = findHeaderForKey(k);
                let v = header ? r[header] : undefined;
                if (k === "sprint" && v !== undefined) {
                  const parsed = parseInt(String(v).match(/\d+/)?.[0] || String(v), 10);
                  v = isNaN(parsed) ? v : parsed;
                } else {
                  v = coerceNumber(v);
                }
                out[k] = v !== undefined ? v : r[k];
              });
              return out;
            });
            // Keep only last MAX_SPRINTS if table has a sprint column
            const hasSprint = expectedKeys.includes("sprint");
            target.rows = hasSprint ? maintainSprintLimit(mapped) : mapped;
          }

          // Recalculate totals/stats for this slide after mapping
          recalculateSlideStats(newSlide);
        } else if (slide.type === "agencyLeads") {
          // Auto-detect which table to update based on headers
          const normalizedHeaders = headers.map((h) => h.toLowerCase().replace(/[^a-z0-9]/g, ""));
          const hasLeadsHeaders = normalizedHeaders.some(
            (h) =>
              h.includes("metrics") ||
              h.includes("totalcount") ||
              h.includes("fromtickets") ||
              h.includes("websiteleads") ||
              h.includes("fromads") ||
              h.includes("livechat") ||
              h.includes("webapp"),
          );
          const hasQ3Headers = normalizedHeaders.some(
            (h) => h.includes("quarter") || h.includes("target") || h.includes("achieved") || h.includes("percentage"),
          );

          const targetKey = hasLeadsHeaders ? "leadsConversion" : hasQ3Headers ? "q3Performance" : "leadsConversion";
          const target = newSlide.data[targetKey];

          if (target?.columns) {
            const expectedKeys: string[] = target.columns.map((c: any) => c.key);
            const mapped = rows.map((r: any) => {
              const out: Record<string, any> = {};
              expectedKeys.forEach((k) => {
                const header = findHeaderForKey(k);
                let v = header ? r[header] : undefined;
                if (k === "sprint" && v !== undefined) {
                  const parsed = parseInt(String(v).match(/\d+/)?.[0] || String(v), 10);
                  v = isNaN(parsed) ? v : parsed;
                } else {
                  v = coerceNumber(v);
                }
                out[k] = v !== undefined ? v : r[k];
              });
              return out;
            });
            target.rows = mapped;
          }

          // Recalculate totals/stats for this slide after mapping
          recalculateSlideStats(newSlide);
        } else if (slide.type === "quarterStats") {
          // Separate sprint rows from quarter stats rows
          const sprintRows = rows.filter((r) => {
            const sprint = String(r.sprint || "").toLowerCase();
            return (
              r.sprint && sprint !== "total" && sprint !== "qtd" && sprint !== "lifetime" && sprint !== "quarter stats"
            );
          });
          if (sprintRows.length > 0) {
            newSlide.data.rows = sprintRows;
          }

          // Update quarter stats if present - look for row without sprint or with special markers
          const statsRow = rows.find((r) => {
            const sprint = String(r.sprint || "").toLowerCase();
            return !r.sprint || sprint === "quarter stats" || sprint === "qtd" || sprint === "total";
          });
          if (statsRow && newSlide.data.quarterStats) {
            Object.keys(newSlide.data.quarterStats).forEach((key) => {
              if (statsRow[key] !== undefined) {
                newSlide.data.quarterStats[key] = statsRow[key];
              }
            });
          }
        } else if (slide.type === "withTarget") {
          // Map incoming headers to expected keys first
          if (newSlide.data?.columns) {
            const expectedKeys: string[] = newSlide.data.columns.map((c: any) => c.key);
            const mapped = rows.map((r: any) => {
              const out: Record<string, any> = {};
              expectedKeys.forEach((k) => {
                const header = findHeaderForKey(k);
                let v = header ? r[header] : undefined;
                if (k === "sprint" && v !== undefined) {
                  const parsed = parseInt(String(v).match(/\d+/)?.[0] || String(v), 10);
                  v = isNaN(parsed) ? v : parsed;
                } else {
                  v = coerceNumber(v);
                }
                out[k] = v !== undefined ? v : r[k];
              });
              return out;
            });

            // Separate sprint rows from totals/qtd rows
            const sprintRows = mapped.filter((r) => {
              const sprint = String(r.sprint).toLowerCase();
              return sprint !== "total" && sprint !== "qtd" && sprint !== "lifetime";
            });
            newSlide.data.rows = sprintRows;

            // Update totals if present
            const totalsRow = mapped.find((r) => {
              const sprint = String(r.sprint).toLowerCase();
              return sprint === "total" || sprint === "qtd";
            });
            if (totalsRow && newSlide.data.total) {
              Object.keys(newSlide.data.total).forEach((key) => {
                if (totalsRow[key] !== undefined) {
                  newSlide.data.total[key] = totalsRow[key];
                }
              });
            }
          }
        } else if (slide.type === "referral") {
          // Map incoming headers to expected keys first
          if (newSlide.data?.columns) {
            const expectedKeys: string[] = newSlide.data.columns.map((c: any) => c.key);
            const mapped = rows.map((r: any) => {
              const out: Record<string, any> = {};
              expectedKeys.forEach((k) => {
                const header = findHeaderForKey(k);
                let v = header ? r[header] : undefined;
                if (k === "sprint" && v !== undefined) {
                  const parsed = parseInt(String(v).match(/\d+/)?.[0] || String(v), 10);
                  v = isNaN(parsed) ? v : parsed;
                } else {
                  v = coerceNumber(v);
                }
                out[k] = v !== undefined ? v : r[k];
              });
              return out;
            });

            // Separate sprint rows from lifetime rows
            const sprintRows = mapped.filter((r) => {
              const sprint = String(r.sprint).toLowerCase();
              return (
                sprint !== "lifetime" &&
                sprint !== "total" &&
                sprint !== "qtd" &&
                (typeof r.sprint === "number" || !isNaN(parseInt(r.sprint)))
              );
            });
            newSlide.data.rows = sprintRows;

            // Update lifetime stats if present
            const lifetimeRow = mapped.find((r) => String(r.sprint).toLowerCase() === "lifetime");
            if (lifetimeRow && newSlide.data.lifetime) {
              Object.keys(newSlide.data.lifetime).forEach((key) => {
                if (lifetimeRow[key] !== undefined) {
                  newSlide.data.lifetime[key] = lifetimeRow[key];
                }
              });
            }
          }
        } else if (slide.type === "wixApp") {
          // Map incoming headers to expected keys first
          if (newSlide.data?.columns) {
            const expectedKeys: string[] = newSlide.data.columns.map((c: any) => c.key);
            const mapped = rows.map((r: any) => {
              const out: Record<string, any> = {};
              expectedKeys.forEach((k) => {
                const header = findHeaderForKey(k);
                let v = header ? r[header] : undefined;
                if (k === "sprint" && v !== undefined) {
                  const parsed = parseInt(String(v).match(/\d+/)?.[0] || String(v), 10);
                  v = isNaN(parsed) ? v : parsed;
                } else {
                  v = coerceNumber(v);
                }
                out[k] = v !== undefined ? v : r[k];
              });
              return out;
            });

            // Separate sprint rows from lifetime/total rows
            const sprintRows = mapped.filter((r) => {
              const sprint = String(r.sprint).toLowerCase();
              return sprint !== "lifetime" && sprint !== "total" && sprint !== "qtd";
            });
            // Keep only last MAX_SPRINTS
            newSlide.data.rows = maintainSprintLimit(sprintRows);

            // Update lifetime stats if present
            const lifetimeRow = mapped.find((r) => String(r.sprint).toLowerCase() === "lifetime");
            if (lifetimeRow && newSlide.data.lifetime) {
              Object.keys(newSlide.data.lifetime).forEach((key) => {
                if (lifetimeRow[key] !== undefined) {
                  newSlide.data.lifetime[key] = lifetimeRow[key];
                }
              });
            }

            // Update total/QTD stats if present
            const totalsRow = mapped.find((r) => {
              const sprint = String(r.sprint).toLowerCase();
              return sprint === "total" || sprint === "qtd";
            });
            if (totalsRow && newSlide.data.total) {
              Object.keys(newSlide.data.total).forEach((key) => {
                if (totalsRow[key] !== undefined) {
                  newSlide.data.total[key] = totalsRow[key];
                }
              });
            }
          }
        } else if (slide.type === "subscriptions") {
          newSlide.data.rows = rows;
        } else if (slide.type === "rankings") {
          // Update position changes table with header mapping and sprint parsing
          const target = newSlide.data.positionChanges;
          if (target?.columns) {
            const expectedKeys: string[] = target.columns.map((c: any) => c.key);
            const mapped = rows.map((r: any) => {
              const out: Record<string, any> = {};
              expectedKeys.forEach((k) => {
                const header = findHeaderForKey(k);
                let v = header ? r[header] : undefined;
                if (k === "sprint" && v !== undefined) {
                  const parsed = parseInt(String(v).match(/\d+/)?.[0] || String(v), 10);
                  v = isNaN(parsed) ? v : parsed;
                } else {
                  v = coerceNumber(v);
                }
                out[k] = v !== undefined ? v : r[k];
              });
              return out;
            });
            // Keep only last MAX_SPRINTS if table has a sprint column
            const hasSprint = expectedKeys.includes("sprint");
            target.rows = hasSprint ? maintainSprintLimit(mapped) : mapped;

            // Recalculate totals/stats for this slide after mapping
            recalculateSlideStats(newSlide);
          }
        } else if (slide.type === "comparison") {
          // Handle comparison slide - separate sprint data from quarter data
          const sprintRows = rows.filter((r) => {
            const type = String(r.type || "").toLowerCase();
            return type === "sprint" || (!r.type && r.sprint && !r.quarter);
          });

          const quarterRows = rows.filter((r) => {
            const type = String(r.type || "").toLowerCase();
            return type === "quarter" || (!r.type && r.quarter);
          });

          if (sprintRows.length > 0) {
            newSlide.data.sprints = sprintRows.map((r: any) => ({
              sprintNumber: coerceNumber(r.sprint || r.sprintnumber),
              paidUsers: coerceNumber(r.paidusers || r.paid),
              totalPaidQTD: coerceNumber(r.totalpaidqtd || r.totalqtd || r.qtd),
            }));
          }

          if (quarterRows.length > 0) {
            newSlide.data.quarters = quarterRows.map((r: any) => ({
              quarter: r.quarter || r.q,
              total: coerceNumber(r.total),
              average: coerceNumber(r.average || r.avg),
            }));
          }
        } else if (slide.type === "pluginRanking") {
          // Handle plugin ranking slide
          const expectedKeys: string[] = newSlide.data.columns?.map((c: any) => c.key) || ["sprint", "position"];
          const mapped = rows.map((r: any) => {
            const out: Record<string, any> = {};
            expectedKeys.forEach((k) => {
              const header = findHeaderForKey(k);
              let v = header ? r[header] : undefined;
              if (k === "sprint" && v !== undefined) {
                const parsed = parseInt(String(v).match(/\d+/)?.[0] || String(v), 10);
                v = isNaN(parsed) ? v : parsed;
              } else {
                v = coerceNumber(v);
              }
              out[k] = v !== undefined ? v : r[k];
            });
            return out;
          });
          // Keep only last MAX_SPRINTS
          newSlide.data.rows = maintainSprintLimit(mapped);
        }

        return newSlide;
      }),
    }));
  };

  const openSlideImport = (slideId) => {
    setCurrentImportSlideId(slideId);
    setPastedData("");
    setPastedImage(null);
    setIsGlobalImport(false);
    setShowImportModal(true);
  };

  const openGlobalImport = () => {
    setCurrentImportSlideId(null);
    setPastedData("");
    setPastedImage(null);
    setIsGlobalImport(true);
    setShowImportModal(true);
  };

  const handleFileImport = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    const fileExtension = file.name.split(".").pop()?.toLowerCase();

    if (fileExtension === "csv") {
      handleCSVImport(file);
    } else if (fileExtension === "json") {
      handleJSONImport(file);
    } else {
      alert("Unsupported file format. Please upload a JSON or CSV file.");
    }

    event.target.value = ""; // Reset input
  };

  const handleJSONImport = (file: File) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const content = e.target?.result as string;
        const importedData = JSON.parse(content);

        if (importedData.version && importedData.slides && Array.isArray(importedData.slides)) {
          // Force a complete state reset with new data
          setSprintData({ slides: [...importedData.slides] });

          if (importedData.currentSprint !== undefined) {
            setCurrentSprint(importedData.currentSprint);
          }

          setShowImportModal(false);
          setIsGlobalImport(false);

          // Force re-render by navigating to first slide
          setCurrentSlide(1);

          alert(
            `✅ Full presentation imported successfully!\n\n${importedData.slides.length} slides loaded\nCurrent Sprint: ${importedData.currentSprint || "Not set"}`,
          );
        } else {
          alert("Invalid import file format. Please use a valid export file.\n\nRequired: version, slides (array)");
        }
      } catch (error) {
        alert(`Error reading import file: ${error.message}\n\nPlease ensure it's a valid JSON export file.`);
      }
    };
    reader.readAsText(file);
  };

  const handleCSVImport = (file: File) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const content = e.target?.result as string;
        const lines = content
          .split("\n")
          .map((l) => l.trim())
          .filter((l) => l && !l.startsWith("#"));

        if (lines.length < 2) {
          alert("CSV file is empty or invalid.");
          return;
        }

        // Parse CSV - handle quoted values properly
        const parseCSVLine = (line: string) => {
          const values: string[] = [];
          let current = "";
          let inQuotes = false;

          for (let i = 0; i < line.length; i++) {
            const char = line[i];
            if (char === '"') {
              inQuotes = !inQuotes;
            } else if (char === "," && !inQuotes) {
              values.push(current.trim());
              current = "";
            } else {
              current += char;
            }
          }
          values.push(current.trim());
          return values;
        };

        const headers = parseCSVLine(lines[0]);
        const rows = lines
          .slice(1)
          .map((line) => {
            const values = parseCSVLine(line);
            const row: any = {};
            headers.forEach((header, idx) => {
              row[header] = values[idx] || "";
            });
            return row;
          })
          .filter((row) => row.SlideID); // Keep rows with SlideID; key column (Sprint/Metrics/Quarter) handled per table

        if (rows.length === 0) {
          alert("No valid data rows found in CSV. Make sure SlideID is filled and headers match your tables.");
          return;
        }

        // Apply the CSV data
        applyMultiSlideCSVData(rows);
      } catch (error) {
        console.error("CSV Import Error:", error);
        alert(`Error reading CSV file: ${error.message}`);
      }
    };
    reader.readAsText(file);
  };

  const applyMultiSlideCSVData = (rows: any[]) => {
    let updatedCount = 0;
    let updatedSlides = new Set<number>();
    const newSlides = [...sprintData.slides];

    rows.forEach((row) => {
      const slideId = parseInt(row.SlideID || row.slideId);
      const slideIndex = newSlides.findIndex((s) => s.id === slideId);

      if (slideIndex === -1) {
        console.warn(`Slide ${slideId} not found, skipping row`);
        return;
      }

      const slide = newSlides[slideIndex];

      try {
        const updated = updateSlideFromCSVRow(slide, row);
        if (updated) {
          updatedCount++;
          updatedSlides.add(slideId);
          // Recalculate totals and stats for the updated slide
          recalculateSlideStats(slide);
        }
      } catch (error) {
        console.error(`Error updating slide ${slideId}:`, error);
      }
    });

    if (updatedCount > 0) {
      setSprintData({ slides: [...newSlides] });
      setShowImportModal(false);
      setIsGlobalImport(false);
      alert(`✅ CSV imported successfully!\n\n${updatedCount} rows updated across ${updatedSlides.size} slides.`);
    } else {
      alert(
        "No data was imported. Please check your CSV format and make sure SlideID, TableName, and Sprint columns are correct.",
      );
    }
  };

  const updateSlideFromCSVRow = (slide: any, row: any): boolean => {
    const tableName = row.TableName || row.tableName || "main";
    const sprintNum = parseInt(row.Sprint);
    // ✅ Ensure table structure exists (fix for Support & Live Chat)
    let targetTable = slide.data[tableName] || slide.data;

    if (!targetTable.rows || !targetTable.columns) {
      targetTable.rows = targetTable.rows || [];
      targetTable.columns = targetTable.columns || Object.keys(row).map((k) => ({ key: k, header: k }));
    }

    if (isNaN(sprintNum)) {
      console.warn(`Invalid sprint number for slide ${slide.id}`);
      return false;
    }

    // Find the target table/data structure

    if (tableName === "main") {
      targetTable = slide.data;
    } else if (slide.data[tableName]) {
      targetTable = slide.data[tableName];
    } else {
      console.warn(`Table "${tableName}" not found in slide ${slide.id}`);
      return false;
    }

    // ✅ Ensure every slide can update, even if structure differs
    if (!targetTable.rows || !targetTable.columns) {
      // Auto-build rows and columns if missing
      targetTable.rows = targetTable.rows || [];
      targetTable.columns =
        targetTable.columns ||
        Object.keys(row).map((key) => ({
          key,
          header: key,
        }));
    }

    // Find or create row for this sprint
    const rowIndex = targetTable.rows.findIndex((r: any) => r.sprint === sprintNum);
    const newRow: any = { sprint: sprintNum };

    // Map CSV columns to table columns
    let hasData = false;
    targetTable.columns.forEach((col: any) => {
      if (col.key === "sprint") return;

      // Try to find matching CSV column
      const csvValue = row[col.header];
      if (csvValue !== undefined && csvValue !== "") {
        hasData = true;
        // Parse value - handle numbers, percentages, times
        const cleanValue = String(csvValue).replace(/[$,%]/g, "").replace(/,/g, "");
        const numValue = parseFloat(cleanValue);
        newRow[col.key] = isNaN(numValue) ? csvValue : numValue;
      }
    });

    if (!hasData) {
      console.warn(`No matching data found for slide ${slide.id}, sprint ${sprintNum}`);
      return false;
    }

    // Update or add the row
    if (rowIndex >= 0) {
      targetTable.rows[rowIndex] = { ...targetTable.rows[rowIndex], ...newRow };
    } else {
      targetTable.rows.push(newRow);
      targetTable.rows = maintainSprintLimit(targetTable.rows);
    }

    return true;
  };

  const downloadSampleCSV = () => {
    // Generate comprehensive CSV with all columns
    const sampleCSV = `SlideID,TableName,Sprint,Position 1-2,Position 3-10,Blog Posts,Infographics,KB Articles,Videos,Total,Social,Blogs,YouTube,Negative,Plugin Position,Total Paid,Direct Plans,Total Tickets Solved,Avg First Response Time,CSAT Score,Pre-sales Tickets,Converted Tickets (Unique customers),Total Paid subscriptions (websites),Agency Tickets,Bad Rating,Conversations assigned,Avg teammate assignment to first response,Total Leads Received,Agency Demos,New Agency Signups,Paid Conversions,Total Count,From Tickets,Website Leads (LP),From Ads,Live chat,Web app (Book a call),Target,Achieved,Percentage,Total Accounts,Paid Trials,Paying Users,Form Fills,Demo,Signups,Paying Agencies,Paid,Shortfall,Trial Signups,Conversions,Referrals,Free Signups,Paid Signups,accountsCreated,cardAdded,bannerActive,payingUsers,roas
1,positionChanges,263,15,45
1,positionChanges,264,18,42
2,main,263,5,3,8,2
2,main,264,6,4,9,3
3,main,263,120,45,30,15,5
3,main,264,125,50,32,18,3
4,main,263,22
4,main,264,21
5,main,263,1500,450
5,main,264,1600,475
7,tickets,263,350,2h 30m,5h 45m,95%,125,8,45,12,Good
7,tickets,264,375,2h 15m,5h 30m,96%,130,10,50,15,Excellent
7,liveChat,263,200,15m,45m,92%,Poor
7,liveChat,264,220,12m,40m,94%,Good
9,main,263,450,85,25
9,main,264,475,90,28
11,main,263,150,35,8
11,main,264,165,38,10
12,main,263,45,12,8,3
12,main,264,50,15,10,4
13,main,263,125,18,67,9
13,main,264,135,22,81,5
14,main,263,25,150,12
14,main,264,30,175,15
15,main,263,15,45,25,8
15,main,264,18,52,30,10
16,main,263,350,25,325,120,15
16,main,264,375,30,345,135,18

# ========================================
# SPRINT DASHBOARD - CSV IMPORT TEMPLATE
# ========================================
#
# INSTRUCTIONS:
# 1. Fill your data in rows below (one row per sprint per table)
# 2. Only fill columns relevant to each slide
# 3. Leave other columns empty
# 4. Import via "Import Data" button in dashboard
#
# COLUMN GUIDE:
# - SlideID: Required - Slide number (see guide below)
# - TableName: Required for slides with tables (see guide below)  
# - Sprint: Required - Sprint number (e.g., 263, 264, 265)
# - Other columns: Match exact column names from your slides
#
# ========================================
# SLIDE REFERENCE GUIDE
# ========================================
#
# Slide 1: Top 25 Major Rankings and Movements
#   TableName: positionChanges
#   Columns: Position 1-2, Position 3-10
#
# Slide 2: Content Publishing Stats
#   TableName: main
#   Columns: Blog Posts, Infographics, KB Articles, Videos
#
# Slide 3: Brand Mentions
#   TableName: main
#   Columns: Total, Social, Blogs, YouTube, Negative
#
# Slide 4: WP Popular Plugin Ranking
#   TableName: main
#   Columns: Plugin Position
#
# Slide 5: Plugin Paid Connections
#   TableName: main
#   Columns: Total Paid, Direct Plans
#
# Slide 7: Support Data (has 2 tables)
#   TableName: tickets
#   Columns: Total Tickets Solved, Avg First Response Time, 
#            CSAT Score, Pre-sales Tickets,
#            Converted Tickets (Unique customers), 
#            Total Paid subscriptions (websites), Agency Tickets, Bad Rating
#   
#   TableName: liveChat
#   Columns: Conversations assigned, Avg teammate assignment to first response, CSAT Score, Bad Rating
#
# Slide 9: Paid Acquisition - Google Ads
#   TableName: main
#   Columns: Total Accounts, Paid Trials, Paying Users
#
# Slide 11: Paid Acquisition - Bing Ads
#   TableName: main
#   Columns: Total Accounts, Paid Trials, Paying Users
#
# Slide 12: Paid Acquisition - Agency Data
#   TableName: main
#   Columns: Form Fills, Demo, Signups, Paying Agencies
#
# Slide 13: Agency - Signups & Paid Users
#   TableName: main
#   Columns: Signups, Paid, Percentage, Shortfall
#
# Slide 14: Partnerships - Affiliate Partner Program
#   TableName: main
#   Columns: Trial Signups, Conversions
#
# Slide 15: Referral Partner Program
#   TableName: main
#   Columns: Trial Signups, Referrals
#
# Slide 16: Strategic Partner Program - Wix App
#   TableName: main
#   Columns: Free Signups, Paid Signups
#
# ========================================
# TIPS:
# - Add multiple rows for same slide (different sprints)
# - Max 5 sprint rows kept per table (oldest auto-removed)
# - Time format: "2h 30m" or "45m"
# - Percentages: "95%" or "95"
# - Leave empty cells blank (don't use 0 unless intentional)
`;

    const blob = new Blob([sampleCSV], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "sprint-dashboard-bulk-import-template.csv";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  const [sprintData, setSprintData] = useState({
    slides: [
      {
        id: 0,
        title: "Sprint Comparison - Paid Users Overview",
        type: "comparison",
        moreDetailsUrl:
          "https://docs.google.com/spreadsheets/d/1O0B4EYLHXCs5s0bWuvlH78cp3WllC2I4MpAw1ifSugU/edit?gid=1339605795#gid=1339605795",
        data: {
          sprints: [
            { sprintNumber: 263, paidUsers: 586, totalPaidQTD: 693 },
            { sprintNumber: 264, paidUsers: 521, totalPaidQTD: 1211 },
            { sprintNumber: 265, paidUsers: 516, totalPaidQTD: 1726 },
            { sprintNumber: 266, paidUsers: 529, totalPaidQTD: 2260 },
          ],
          quarters: [
            { quarter: "Q1", total: 5278, average: 468 },
            { quarter: "Q2", total: 6247, average: 475 },
            { quarter: "Q3", total: 0, average: 0 },
          ],
        },
      },
      {
        id: 2,
        title: "Content Publishing Stats",
        type: "table",
        moreDetailsUrl:
          "https://docs.google.com/spreadsheets/d/1O0B4EYLHXCs5s0bWuvlH78cp3WllC2I4MpAw1ifSugU/edit?gid=1006810399#gid=1006810399",
        data: {
          columns: [
            { key: "sprint", header: "Sprint", locked: true },
            { key: "blog", header: "Blog Posts" },
            { key: "infographics", header: "Infographics" },
            { key: "kb", header: "KB Articles" },
            { key: "videos", header: "Videos" },
          ],
          rows: [
            { sprint: 263, blog: 0, infographics: 0, kb: 0, videos: 0 },
            { sprint: 264, blog: 2, infographics: 0, kb: 0, videos: 1 },
            { sprint: 265, blog: 1, infographics: 0, kb: 1, videos: 1 },
            { sprint: 266, blog: 2, infographics: 0, kb: 6, videos: 0 },
          ],
        },
      },
      {
        id: 1,
        title: "Top 25 Major Rankings and Movements",
        type: "rankings",
        moreDetailsUrl:
          "https://docs.google.com/spreadsheets/d/1O0B4EYLHXCs5s0bWuvlH78cp3WllC2I4MpAw1ifSugU/edit?gid=769589025#gid=769589025",
        data: {
          total: 0,
          byRegion: [
            { region: "US", count: 0, keywords: "" },
            { region: "UK", count: 0, keywords: "" },
            { region: "DE", count: 0, keywords: "" },
          ],
          positionChanges: {
            columns: [
              { key: "sprint", header: "Sprint", locked: true },
              { key: "pos1_2", header: "Position 1-2" },
              { key: "pos3_10", header: "Position 3-10" },
            ],
            rows: [
              { sprint: 263, pos1_2: 15, pos3_10: 45 },
              { sprint: 264, pos1_2: 18, pos3_10: 42 },
              { sprint: 265, pos1_2: 15, pos3_10: 49 },
              { sprint: 266, pos1_2: 17, pos3_10: 47 },
            ],
          },
          improved: [{ region: "US", count: 0, keywords: "" }],
          declined: [{ region: "US", count: 0, keywords: "" }],
        },
      },
      {
        id: 3,
        title: "Brand Mentions",
        type: "table",
        moreDetailsUrl:
          "https://docs.google.com/spreadsheets/d/1O0B4EYLHXCs5s0bWuvlH78cp3WllC2I4MpAw1ifSugU/edit?gid=1091728359#gid=1091728359",
        data: {
          columns: [
            { key: "sprint", header: "Sprint", locked: true },
            { key: "total", header: "Total" },
            { key: "social", header: "Social" },
            { key: "blogs", header: "Blogs" },
            { key: "youtube", header: "YouTube" },
            { key: "negative", header: "Negative" },
          ],
          rows: [
            { sprint: 263, total: 39, social: 13, blogs: 6, youtube: 0, negative: 3 },
            { sprint: 264, total: 99, social: 9, blogs: 8, youtube: 1, negative: 1 },
            { sprint: 265, total: 135, social: 8, blogs: 4, youtube: 1, negative: 0 },
            { sprint: 266, total: 62, social: 7, blogs: 4, youtube: 0, negative: 0 },
          ],
        },
      },
      {
        id: 4,
        title: "WP Popular Plugin Ranking",
        type: "pluginRanking",
        moreDetailsUrl:
          "https://docs.google.com/spreadsheets/d/1O0B4EYLHXCs5s0bWuvlH78cp3WllC2I4MpAw1ifSugU/edit?gid=1464999886#gid=1464999886",
        data: {
          quarterTarget: 21,
          columns: [
            { key: "sprint", header: "Sprint", locked: true },
            { key: "position", header: "Plugin Position" },
          ],
          rows: [
            { sprint: 263, position: 39 },
            { sprint: 264, position: 39 },
            { sprint: 265, position: 39 },
            { sprint: 266, position: 39 },
          ],
        },
      },
      {
        id: 5,
        title: "Plugin Paid Connections",
        type: "table",
        moreDetailsUrl:
          "https://docs.google.com/spreadsheets/d/1O0B4EYLHXCs5s0bWuvlH78cp3WllC2I4MpAw1ifSugU/edit?gid=929138315#gid=929138315",
        data: {
          columns: [
            { key: "sprint", header: "Sprint", locked: true },
            { key: "total", header: "Total Paid" },
            { key: "direct", header: "Direct Plans" },
          ],
          rows: [
            { sprint: 263, total: 142, direct: 96 },
            { sprint: 264, total: 87, direct: 52 },
            { sprint: 265, total: 82, direct: 61 },
            { sprint: 266, total: 83, direct: 55 },
          ],
        },
      },
      {
        id: 7,
        title: "Support data",
        type: "supportData",
        moreDetailsUrl:
          "https://docs.google.com/spreadsheets/d/1O0B4EYLHXCs5s0bWuvlH78cp3WllC2I4MpAw1ifSugU/edit?gid=40331932#gid=40331932",
        data: {
          tickets: {
            columns: [
              { key: "sprint", header: "Sprint", locked: true },
              { key: "totalTickets", header: "Total Tickets Solved" },
              { key: "avgFirstResponse", header: "Avg First Response Time" },
              { key: "csat", header: "CSAT Score" },
              { key: "presales", header: "Pre-sales Tickets" },
              { key: "converted", header: "Converted Tickets (Unique customers)" },
              { key: "paidSubs", header: "Total Paid subscriptions (websites)" },
              { key: "agencyTickets", header: "Agency Tickets" },
              { key: "badRating", header: "Bad Rating" },
            ],
            rows: [
              {
  sprint: 264,
  totalTickets: 56,
  avgFirstResponse: "5 hrs",
  csat: "75% (3:1)",
  presales: 19,
  converted: 4,
  paidSubs: 12,
  agencyTickets: 12,
  badRating: "Asking for promo code",
},
{
  sprint: 265,
  totalTickets: 47,
  avgFirstResponse: "7.6 hrs",
  csat: "100% (4:0)",
  presales: 8,
  converted: 2,
  paidSubs: 2,
  agencyTickets: 12,
  badRating: "-",
},
{
  sprint: 266,
  totalTickets: 51,
  avgFirstResponse: "8.4 hrs",
  csat: "100% (2:0)",
  presales: 23,
  converted: 5,
  paidSubs: 5,
  agencyTickets: 11,
  badRating: "-",
},
            ],
          },
          liveChat: {
            columns: [
              { key: "sprint", header: "Sprint", locked: true },
              { key: "conversations", header: "Conversations assigned" },
              { key: "avgAssignment", header: "Avg teammate assignment to first response" },
              { key: "csat", header: "CSAT Score" },
              { key: "badRating", header: "Bad Rating" },
            ],
            rows: [
              { sprint: 264, conversations: 2, avgAssignment: "28s", csat: "-", badRating: "-" },
              { sprint: 265, conversations: "-", avgAssignment: "-", csat: "-", badRating: "-" },
              { sprint: 266, conversations: 3, avgAssignment: "1 hr 19 m", csat: "100% (2:0)", badRating: "-" },
            ],
          },
        },
      },
      {
        id: 8,
        title: "Agency Leads & Conversion (Presales)",
        type: "agencyLeads",
        moreDetailsUrl:
          "https://docs.google.com/spreadsheets/d/1O0B4EYLHXCs5s0bWuvlH78cp3WllC2I4MpAw1ifSugU/edit?gid=479602000#gid=479602000",
        data: {
          leadsConversion: {
            columns: [
              { key: "metrics", header: "Metrics", locked: true },
              { key: "totalCount", header: "Total Count" },
              { key: "fromTickets", header: "From Tickets" },
              { key: "websiteLeads", header: "Website Leads (LP)" },
              { key: "fromAds", header: "From Ads" },
              { key: "liveChat", header: "Live chat" },
              { key: "webApp", header: "Web app (Book a call)" },
            ],
            rows: [
              {
                metrics: "Total Leads Received",
                totalCount: 0,
                fromTickets: 0,
                websiteLeads: 0,
                fromAds: 0,
                liveChat: 0,
                webApp: 0,
              },
              {
                metrics: "Agency Demos",
                totalCount: 0,
                fromTickets: 0,
                websiteLeads: 0,
                fromAds: 0,
                liveChat: 0,
                webApp: 0,
              },
              {
                metrics: "New Agency Signups",
                totalCount: 0,
                fromTickets: 0,
                websiteLeads: 0,
                fromAds: 0,
                liveChat: 0,
                webApp: 0,
              },
              {
                metrics: "Paid Conversions",
                totalCount: 0,
                fromTickets: 0,
                websiteLeads: 0,
                fromAds: 0,
                liveChat: 0,
                webApp: 0,
              },
            ],
          },
          q3Performance: {
            columns: [
              { key: "quarter", header: "Quarter", locked: true },
              { key: "target", header: "Target" },
              { key: "achieved", header: "Achieved" },
              { key: "percentage", header: "Percentage" },
            ],
            rows: [
              { quarter: "Sprint 263", target: 0, achieved: 0, percentage: "0%" },
              { quarter: "Sprint 264", target: 0, achieved: 0, percentage: "0%" },
              { quarter: "Sprint 265", target: 0, achieved: 0, percentage: "0%" },
            ],
          },
        },
      },
      {
        id: 9,
        title: "Paid Acquisition - Google Ads",
        type: "quarterStats",
        moreDetailsUrl:
          "https://docs.google.com/spreadsheets/d/1O0B4EYLHXCs5s0bWuvlH78cp3WllC2I4MpAw1ifSugU/edit?gid=1295636005#gid=1295636005",
        data: {
          quarterStats: {
            accountsCreated: 0,
            cardAdded: 0,
            bannerActive: 0,
            payingUsers: 0,
            roas: 0,
          },
          columns: [
            { key: "sprint", header: "Sprint", locked: true },
            { key: "totalAccounts", header: "Total Accounts" },
            { key: "paidTrials", header: "Paid Trials" },
            { key: "payingUsers", header: "Paying Users" },
          ],
          rows: [
            { sprint: 263, totalAccounts: 135, paidTrials: 134, payingUsers: 36 },
            { sprint: 264, totalAccounts: 181, paidTrials: 180, payingUsers: 38 },
            { sprint: 265, totalAccounts: 139, paidTrials: 135, payingUsers: 25 },
          ],
        },
      },
      {
        id: 10,
        title: "Google Ads Key Observations",
        type: "googleAdsObservations",
        moreDetailsUrl: "https://docs.google.com/spreadsheets",
        data: {
          performance: {
            impressions: { value: 0, isIncrease: true },
            clicks: { value: 0, isIncrease: true },
            CTR: { value: 0, isIncrease: true },
            conversions: { value: 0, isIncrease: false },
          },
          auctionInsights: {
            impressionShare: { value: 0, isIncrease: true },
            absoluteTopOfPage: { value: 0, isIncrease: false },
          },
        },
      },
      {
        id: 11,
        title: "Paid Acquisition - Bing Ads",
        type: "quarterStats",
        moreDetailsUrl:
          "https://docs.google.com/spreadsheets/d/1O0B4EYLHXCs5s0bWuvlH78cp3WllC2I4MpAw1ifSugU/edit?gid=1295636005#gid=1295636005",
        data: {
          quarterStats: {
            accountsCreated: 0,
            cardAdded: 0,
            bannerActive: 0,
            payingUsers: 0,
            roas: 0,
          },
          columns: [
            { key: "sprint", header: "Sprint", locked: true },
            { key: "totalAccounts", header: "Total Accounts" },
            { key: "paidTrials", header: "Paid Trials" },
            { key: "payingUsers", header: "Paying Users" },
          ],
          rows: [
            { sprint: 263, totalAccounts: 10, paidTrials: 9, payingUsers: 5 },
            { sprint: 264, totalAccounts: 10, paidTrials: 9, payingUsers: 0 },
            { sprint: 265, totalAccounts: 5, paidTrials: 5, payingUsers: 2 },
          ],
        },
      },
      {
        id: 12,
        title: "Paid Acquisition - Agency Data",
        type: "table",
        moreDetailsUrl:
          "https://docs.google.com/spreadsheets/d/1O0B4EYLHXCs5s0bWuvlH78cp3WllC2I4MpAw1ifSugU/edit?gid=1295636005#gid=1295636005",
        data: {
          columns: [
            { key: "sprint", header: "Sprint", locked: true },
            { key: "formFills", header: "Form Fills" },
            { key: "demos", header: "Demo" },
            { key: "signups", header: "Signups" },
            { key: "paying", header: "Paying Agencies" },
          ],
          rows: [
            { sprint: 263, formFills: 16, demos: 2, signups: 0, paying: 0 },
            { sprint: 264, formFills: 17, demos: 0, signups: 1, paying: 0 },
            { sprint: 265, formFills: 14, demos: 0, signups: 4, paying: 0 },
            { sprint: 265, formFills: 8, demos: 0, signups: 0, paying: 0 },
          ],
        },
      },
      {
        id: 13,
        title: "Agency - Signups & Paid Users",
        type: "withTarget",
        moreDetailsUrl:
          "https://docs.google.com/spreadsheets/d/1O0B4EYLHXCs5s0bWuvlH78cp3WllC2I4MpAw1ifSugU/edit?gid=1941259614#gid=1941259614",
        data: {
          target: 27,
          targetLabel: "SPRINT PAID TARGET",
          hideQtdStats: true,
          columns: [
            { key: "sprint", header: "Sprint", locked: true },
            { key: "signups", header: "Signups" },
            { key: "paid", header: "Paid" },
            { key: "percentage", header: "Target Achieved %" },
            { key: "shortfall", header: "Shortfall" },
          ],
          rows: [
           { sprint: 263, signups: 34, paid: 8, percentage: 30, shortfall: 19 },
            { sprint: 264, signups: 28, paid: 8, percentage: 30, shortfall: 19 },
            { sprint: 265, signups: 26, paid: 6, percentage: 22, shortfall: 21 },
            { sprint: 266, signups: 23, paid: 7, percentage: 26, shortfall: 20 },
          ],
          total: { signups: 0, paid: 0 },
        },
      },
      {
        id: 14,
        title: "Partnerships & Growth - Affiliate Partner Program",
        type: "withTarget",
        moreDetailsUrl:
          "https://docs.google.com/spreadsheets/d/1O0B4EYLHXCs5s0bWuvlH78cp3WllC2I4MpAw1ifSugU/edit?gid=801337408#gid=801337408",
        data: {
          target: 1000,
          targetLabel: "QUARTER TARGET",
          columns: [
            { key: "sprint", header: "Sprint", locked: true },
            { key: "trials", header: "Trial Signups" },
            { key: "conversions", header: "Conversions" },
          ],
          rows: [
            { sprint: 262, trials: 11, conversions: 24 },
            { sprint: 263, trials: 6, conversions: 35 },
            { sprint: 264, trials: 11, conversions: 27 },
            { sprint: 265, trials: 27, conversions: 13 },
            { sprint: 266, trials: 17, conversions: 23 },
          ],
          total: { trials: 72, conversions: 122 },
        },
      },
      {
        id: 16,
        title: "Partnerships & Growth - Strategic Partner Program - Wix App",
        type: "wixApp",
        moreDetailsUrl:
          "https://docs.google.com/spreadsheets/d/1O0B4EYLHXCs5s0bWuvlH78cp3WllC2I4MpAw1ifSugU/edit?gid=801337408#gid=801337408",
        data: {
          lifetime: { paid: 2106, rating: 3.43 },
          target: 400,
          current: 61,
          columns: [
            { key: "sprint", header: "Sprint", locked: true },
            { key: "freeSignups", header: "Free Signups" },
            { key: "paid", header: "Paid Signups" },
          ],
          rows: [
            { sprint: 262, freeSignups: 90, paid: 12 },
            { sprint: 263, freeSignups: 97, paid: 9 },
            { sprint: 264, freeSignups: 90, paid: 16 },
            { sprint: 265, freeSignups: 98, paid: 16 },
            { sprint: 266, freeSignups: 88, paid: 8 },
          ],
          total: { freeSignups: 463, paid: 61 },
        },
      },
      {
        id: 15,
        title: "Partnerships & Growth - Referral Partner Program",
        type: "referral",
        moreDetailsUrl:
          "https://docs.google.com/spreadsheets/d/1O0B4EYLHXCs5s0bWuvlH78cp3WllC2I4MpAw1ifSugU/edit?gid=801337408#gid=801337408",
        data: {
          lifetime: { paid: 921 },
          target: 100,
          current: 0,
          columns: [
            { key: "sprint", header: "Sprint", locked: true },
            { key: "trials", header: "Trial Signups" },
            { key: "referrals", header: "Referrals" },
          ],
          rows: [
            { sprint: 262, trials: 0, referrals: 6 },
            { sprint: 263, trials: 1, referrals: 2 },
            { sprint: 264, trials: 4, referrals: 4 },
            { sprint: 265, trials: 5, referrals: 0 },
            { sprint: 266, trials: 3, referrals: 1 },
          ],
          total: { trials: 13, referrals: 13 },
        },
      },
      {
        id: 17,
        title: "New subscriptions & paid signups",
        type: "subscriptions",
        moreDetailsUrl:
          "https://docs.google.com/spreadsheets/d/1O0B4EYLHXCs5s0bWuvlH78cp3WllC2I4MpAw1ifSugU/edit?gid=458351126#gid=458351126",
        data: {
          rows: [
            { channel: "Direct (new subscription)", totalTarget: 11009, targetAsOnDate: 0, actual: 0, percentage: 0 },
            { channel: "Agency (new subscription)", totalTarget: 315, targetAsOnDate: 0, actual: 0, percentage: 0 },
            { channel: "Affiliate (paid signups)", totalTarget: 1000, targetAsOnDate: 0, actual: 0, percentage: 0 },
            { channel: "Ads", totalTarget: 800, targetAsOnDate: 0, actual: 0, percentage: 0 },
          ],
        },
      },
    ],
  });

  const updateSlideTitle = (slideId, newTitle) => {
    preserveScroll();
    setSprintData((prev) => ({
      ...prev,
      slides: prev.slides.map((s) => (s.id === slideId ? { ...s, title: newTitle } : s)),
    }));
  };

  const updateMoreDetailsUrl = (slideId, url) => {
    preserveScroll();
    setSprintData((prev) => ({
      ...prev,
      slides: prev.slides.map((s) => (s.id === slideId ? { ...s, moreDetailsUrl: url } : s)),
    }));
  };

  const updateSlideData = (slideId, path, value) => {
    preserveScroll();
    setSprintData((prev) => ({
      ...prev,
      slides: prev.slides.map((s) => {
        const actualId = getActualSlideId(slideId);

        if (s.id !== actualId) return s;
        const newSlide = JSON.parse(JSON.stringify(s));
        let current = newSlide.data;

        if (s.type === "supportData") {
          current = slideId > 1000 ? current.liveChat : current.tickets;
        } else if (s.type === "agencyLeads") {
          current = slideId > 2000 ? current.q3Performance : current.leadsConversion;
        } else if (s.type === "rankings" && slideId === s.id && path[0] === "rows") {
          current = current.positionChanges;
        }

        for (let i = 0; i < path.length - 1; i++) {
          current = current[path[i]];
        }
        current[path[path.length - 1]] = isNaN(value) ? value : Number(value);
        return newSlide;
      }),
    }));
  };

  const addRow = (slideId) => {
    preserveScroll();
    setSprintData((prev) => ({
      ...prev,
      slides: prev.slides.map((s) => {
        const actualId = getActualSlideId(slideId);

        if (s.id !== actualId) return s;
        const newSlide = JSON.parse(JSON.stringify(s));

        let targetData = newSlide.data;
        if (s.type === "supportData") {
          targetData = slideId > 1000 ? newSlide.data.liveChat : newSlide.data.tickets;
        } else if (s.type === "agencyLeads") {
          targetData = slideId > 2000 ? newSlide.data.q3Performance : newSlide.data.leadsConversion;
        } else if (s.type === "rankings" && slideId === s.id) {
          targetData = newSlide.data.positionChanges;
        }

        const rows = targetData.rows;
        if (rows && rows.length > 0) {
          const lastRow = rows[rows.length - 1];
          const newRow: any = {};
          targetData.columns.forEach((col) => {
            if (col.key === "sprint" || col.key === "metrics" || col.key === "quarter") {
              if (col.key === "sprint" && typeof lastRow.sprint === "number") {
                newRow[col.key] = lastRow.sprint + 1;
              } else {
                newRow[col.key] = "";
              }
            } else {
              newRow[col.key] = typeof lastRow[col.key] === "number" ? 0 : "";
            }
          });
          rows.push(newRow);

          // Maintain sprint limit
          if (targetData.columns.some((c) => c.key === "sprint")) {
            targetData.rows = maintainSprintLimit(rows);
          }

          // Recalculate stats for all table types
          recalculateSlideStats(newSlide);
        }
        return newSlide;
      }),
    }));
  };

  const removeRow = (slideId, rowIndex) => {
    preserveScroll();
    setSprintData((prev) => ({
      ...prev,
      slides: prev.slides.map((s) => {
        const actualId = getActualSlideId(slideId);

        if (s.id !== actualId) return s;
        const newSlide = JSON.parse(JSON.stringify(s));

        let targetData = newSlide.data;
        if (s.type === "supportData") {
          targetData = slideId > 1000 ? newSlide.data.liveChat : newSlide.data.tickets;
        } else if (s.type === "agencyLeads") {
          targetData = slideId > 2000 ? newSlide.data.q3Performance : newSlide.data.leadsConversion;
        } else if (s.type === "rankings" && slideId === s.id) {
          targetData = newSlide.data.positionChanges;
        }

        if (targetData.rows && targetData.rows.length > 1) {
          targetData.rows.splice(rowIndex, 1);

          // Recalculate stats for all table types
          recalculateSlideStats(newSlide);
        } else {
          alert("Cannot delete the last row. At least one row is required.");
        }
        return newSlide;
      }),
    }));
  };

  const addColumn = (slideId) => {
    preserveScroll();
    setSprintData((prev) => ({
      ...prev,
      slides: prev.slides.map((s) => {
        const actualId = getActualSlideId(slideId);

        if (s.id !== actualId) return s;
        const newSlide = JSON.parse(JSON.stringify(s));

        let targetData = newSlide.data;
        if (s.type === "supportData") {
          targetData = slideId > 1000 ? newSlide.data.liveChat : newSlide.data.tickets;
        } else if (s.type === "agencyLeads") {
          targetData = slideId > 2000 ? newSlide.data.q3Performance : newSlide.data.leadsConversion;
        } else if (s.type === "rankings" && slideId === s.id) {
          targetData = newSlide.data.positionChanges;
        }

        const newColKey = `col${Date.now()}`;
        targetData.columns.push({ key: newColKey, header: "New Column" });
        targetData.rows.forEach((row) => {
          row[newColKey] = 0;
        });
        return newSlide;
      }),
    }));
  };

  const removeColumn = (slideId, colKey) => {
    preserveScroll();
    setSprintData((prev) => ({
      ...prev,
      slides: prev.slides.map((s) => {
        const actualId = getActualSlideId(slideId);

        if (s.id !== actualId) return s;
        const newSlide = JSON.parse(JSON.stringify(s));

        let targetData = newSlide.data;
        if (s.type === "supportData") {
          targetData = slideId > 1000 ? newSlide.data.liveChat : newSlide.data.tickets;
        } else if (s.type === "agencyLeads") {
          targetData = slideId > 2000 ? newSlide.data.q3Performance : newSlide.data.leadsConversion;
        } else if (s.type === "rankings" && slideId === s.id) {
          targetData = newSlide.data.positionChanges;
        }

        const nonLockedColumns = targetData.columns.filter((c) => !c.locked);
        if (nonLockedColumns.length <= 1) {
          alert("Cannot delete the last non-locked column. At least one data column is required.");
          return s;
        }

        targetData.columns = targetData.columns.filter((c) => c.key !== colKey);
        targetData.rows.forEach((row) => {
          delete row[colKey];
        });
        return newSlide;
      }),
    }));
  };

  const updateColumnHeader = (slideId, colKey, newHeader) => {
    preserveScroll();
    setSprintData((prev) => ({
      ...prev,
      slides: prev.slides.map((s) => {
        const actualId = getActualSlideId(slideId);

        if (s.id !== actualId) return s;
        const newSlide = JSON.parse(JSON.stringify(s));

        let targetData = newSlide.data;
        if (s.type === "supportData") {
          targetData = slideId > 1000 ? newSlide.data.liveChat : newSlide.data.tickets;
        } else if (s.type === "agencyLeads") {
          targetData = slideId > 2000 ? newSlide.data.q3Performance : newSlide.data.leadsConversion;
        } else if (s.type === "rankings" && slideId === s.id) {
          targetData = newSlide.data.positionChanges;
        }

        const col = targetData.columns.find((c) => c.key === colKey);
        if (col) col.header = newHeader;
        return newSlide;
      }),
    }));
  };

  const updateCellValue = (slideId, rowIndex, colKey, newValue) => {
    preserveScroll();
    setSprintData((prev) => ({
      ...prev,
      slides: prev.slides.map((s) => {
        const actualId = getActualSlideId(slideId);

        if (s.id !== actualId) return s;
        const newSlide = JSON.parse(JSON.stringify(s));

        let targetData = newSlide.data;
        if (s.type === "supportData") {
          targetData = slideId > 1000 ? newSlide.data.liveChat : newSlide.data.tickets;
        } else if (s.type === "agencyLeads") {
          targetData = slideId > 2000 ? newSlide.data.q3Performance : newSlide.data.leadsConversion;
        } else if (s.type === "rankings" && slideId === s.id) {
          targetData = newSlide.data.positionChanges;
        }

        if (targetData.rows && targetData.rows[rowIndex]) {
          const parsedValue = isNaN(newValue) || newValue === "" ? newValue : Number(newValue);
          targetData.rows[rowIndex][colKey] = parsedValue;

          // Recalculate stats for all table types
          recalculateSlideStats(newSlide);
        }
        return newSlide;
      }),
    }));
  };

  const deleteRow = (slideId, rowIdx) => {
    preserveScroll();
    setSprintData((prev) => ({
      ...prev,
      slides: prev.slides.map((s) => {
        const actualId = getActualSlideId(slideId);

        if (s.id !== actualId) return s;
        const newSlide = JSON.parse(JSON.stringify(s));

        let targetData = newSlide.data;
        if (s.type === "supportData") {
          targetData = slideId > 1000 ? newSlide.data.liveChat : newSlide.data.tickets;
        } else if (s.type === "agencyLeads") {
          targetData = slideId > 2000 ? newSlide.data.q3Performance : newSlide.data.leadsConversion;
        } else if (s.type === "rankings" && slideId === s.id) {
          targetData = newSlide.data.positionChanges;
        }

        if (targetData.rows && targetData.rows.length > 1) {
          targetData.rows.splice(rowIdx, 1);

          // Recalculate stats for all table types
          recalculateSlideStats(newSlide);
        }

        return newSlide;
      }),
    }));
  };

  const deleteColumn = (slideId, colKey) => {
    preserveScroll();
    setSprintData((prev) => ({
      ...prev,
      slides: prev.slides.map((s) => {
        const actualId = getActualSlideId(slideId);

        if (s.id !== actualId) return s;
        const newSlide = JSON.parse(JSON.stringify(s));

        let targetData = newSlide.data;
        if (s.type === "supportData") {
          targetData = slideId > 1000 ? newSlide.data.liveChat : newSlide.data.tickets;
        } else if (s.type === "agencyLeads") {
          targetData = slideId > 2000 ? newSlide.data.q3Performance : newSlide.data.leadsConversion;
        } else if (s.type === "rankings" && slideId === s.id) {
          targetData = newSlide.data.positionChanges;
        }

        const nonLockedColumns = targetData.columns.filter((c) => !c.locked);
        if (nonLockedColumns.length <= 1) {
          alert("Cannot delete the last non-locked column. At least one data column is required.");
          return s;
        }

        targetData.columns = targetData.columns.filter((c) => c.key !== colKey);
        targetData.rows.forEach((row) => {
          delete row[colKey];
        });
        return newSlide;
      }),
    }));
  };

  const exportToJSON = () => {
    try {
      const exportData = {
        version: "1.0",
        exportDate: new Date().toISOString(),
        currentSprint,
        slides: sprintData.slides,
      };

      const dataStr = JSON.stringify(exportData, null, 2);
      const dataBlob = new Blob([dataStr], { type: "application/json" });
      const url = URL.createObjectURL(dataBlob);
      const link = document.createElement("a");
      link.href = url;
      link.download = `sprint-dashboard-export-${currentSprint}-${new Date().toISOString().split("T")[0]}.json`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);

      alert("✅ Data exported successfully! You can use this file to import and update your presentation later.");
    } catch (error) {
      console.error("Export error:", error);
      alert("An error occurred while exporting data. Please try again.");
    }
  };

  const exportToPDF = () => {
    if (isEditMode) {
      alert("Please save your changes (exit Edit mode) before exporting to PDF");
      return;
    }

    // Add a small delay to ensure any state changes are rendered
    setTimeout(() => {
      try {
        window.print();
      } catch (error) {
        console.error("Print failed:", error);
        alert("Unable to open print dialog. Please try using Ctrl+P (Windows) or Cmd+P (Mac) instead.");
      }
    }, 100);
  };

  const enterPresentMode = () => {
    setIsPresentMode(true);
    setCurrentSlide(1);
    const elem = document.documentElement;
    if (elem.requestFullscreen) {
      elem.requestFullscreen().catch((err) => {
        console.log("Fullscreen request failed:", err);
      });
    }
  };

  const exitPresentMode = () => {
    setIsPresentMode(false);
    if (document.fullscreenElement) {
      document.exitFullscreen();
    }
  };

  const goToNextSlide = () => {
    if (currentSlide < sprintData.slides.length) {
      setCurrentSlide(currentSlide + 1);
    }
  };

  const goToPreviousSlide = () => {
    if (currentSlide > 1) {
      setCurrentSlide(currentSlide - 1);
    }
  };

  useEffect(() => {
    if (!isPresentMode) return;

    const handleKeyDown = (e) => {
      if (e.key === "ArrowRight" || e.key === "ArrowDown" || e.key === " ") {
        e.preventDefault();
        goToNextSlide();
      } else if (e.key === "ArrowLeft" || e.key === "ArrowUp") {
        e.preventDefault();
        goToPreviousSlide();
      } else if (e.key === "Escape") {
        e.preventDefault();
        exitPresentMode();
      } else if (e.key === "Home") {
        e.preventDefault();
        setCurrentSlide(1);
      } else if (e.key === "End") {
        e.preventDefault();
        setCurrentSlide(sprintData.slides.length);
      }
    };

    window.addEventListener("keydown", handleKeyDown);
    return () => window.removeEventListener("keydown", handleKeyDown);
  }, [isPresentMode, currentSlide, sprintData.slides.length]);

  const EditableTable = ({ slide }) => {
    const data = slide.data;
    if (!data || !data.rows || !data.columns) return null;
    const lastIdx = data.rows.length - 1;
    const actualIdForEdit = isNestedTableSlide(slide.type) ? getActualSlideId(slide.id) : slide.id;
    const isEditing = editingTableId === slide.id || editingTableId === actualIdForEdit;

    return (
      <div style={{ marginBottom: "12px" }}>
        {isEditMode && (
          <div style={{ marginBottom: "12px", display: "flex", gap: "8px", flexWrap: "wrap", alignItems: "center" }}>
            {!isEditing && (
              <button
                onClick={() => setEditingTableId(slide.id)}
                style={{
                  padding: "8px 16px",
                  fontSize: "13px",
                  fontWeight: "600",
                  backgroundColor: "#1863DC",
                  color: "#fff",
                  border: "none",
                  borderRadius: "20px",
                  cursor: "pointer",
                  whiteSpace: "nowrap",
                }}
              >
                ✏️ Edit Table
              </button>
            )}
            {isEditing && (
              <>
                <button
                  onClick={() => setEditingTableId(null)}
                  style={{
                    padding: "8px 16px",
                    fontSize: "13px",
                    fontWeight: "600",
                    backgroundColor: "#2DAD70",
                    color: "#fff",
                    border: "none",
                    borderRadius: "20px",
                    cursor: "pointer",
                    whiteSpace: "nowrap",
                  }}
                >
                  ✓ Done
                </button>
                <button
                  onClick={() => addRow(slide.id)}
                  style={{
                    padding: "8px 16px",
                    fontSize: "13px",
                    fontWeight: "600",
                    backgroundColor: "#7F56D9",
                    color: "#fff",
                    border: "none",
                    borderRadius: "20px",
                    cursor: "pointer",
                    whiteSpace: "nowrap",
                  }}
                >
                  + Add Row
                </button>
                <button
                  onClick={() => addColumn(slide.id)}
                  style={{
                    padding: "8px 16px",
                    fontSize: "13px",
                    fontWeight: "600",
                    backgroundColor: "#FF9A3C",
                    color: "#fff",
                    border: "none",
                    borderRadius: "20px",
                    cursor: "pointer",
                    whiteSpace: "nowrap",
                  }}
                >
                  + Add Column
                </button>
              </>
            )}
            <button
              onClick={() => {
                // For nested tables, use the actual parent slide ID
                const parentSlideId = isNestedTableSlide(slide.type) ? getActualSlideId(slide.id) : slide.id;
                openSlideImport(parentSlideId);
              }}
              style={{
                padding: "8px 16px",
                fontSize: "13px",
                fontWeight: "600",
                backgroundColor: "#2DAD70",
                color: "#fff",
                border: "none",
                borderRadius: "20px",
                cursor: "pointer",
                whiteSpace: "nowrap",
              }}
            >
              🔄 Update Table Data
            </button>
          </div>
        )}
        <div style={{ overflowX: "auto" }}>
          <table
            style={{
              width: "100%",
              borderCollapse: "collapse",
              fontSize: "14px",
              fontFamily: "Inter, -apple-system, BlinkMacSystemFont, sans-serif",
            }}
          >
            <thead>
              <tr style={{ backgroundColor: "#F8FAFB", borderBottom: "2px solid #1863DC" }}>
                {isEditing && (
                  <th
                    style={{
                      padding: "12px 8px",
                      textAlign: "center",
                      fontWeight: "600",
                      color: "#212121",
                      width: "40px",
                    }}
                  ></th>
                )}
                {data.columns.map((col) => (
                  <th
                    key={col.key}
                    style={{ padding: "12px 16px", textAlign: "left", fontWeight: "600", color: "#212121" }}
                  >
                    <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
                      {isEditing && !col.locked ? (
                        <input
                          type="text"
                          defaultValue={col.header}
                          onBlur={(e) => updateColumnHeader(slide.id, col.key, e.target.value)}
                          style={{
                            width: "100%",
                            padding: "4px",
                            border: "1px solid #1863DC",
                            borderRadius: "4px",
                            fontWeight: "600",
                          }}
                        />
                      ) : (
                        col.header
                      )}
                      {isEditing && !col.locked && data.columns.length > 1 && (
                        <button
                          onClick={() => deleteColumn(slide.id, col.key)}
                          style={{
                            background: "transparent",
                            border: "none",
                            cursor: "pointer",
                            fontSize: "16px",
                            padding: "4px",
                            color: "#DC2143",
                          }}
                          title="Delete column"
                        >
                          🗑️
                        </button>
                      )}
                    </div>
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {data.rows.map((row, rowIdx) => (
                <tr
                  key={rowIdx}
                  style={{
                    backgroundColor: rowIdx % 2 ? "#F8FAFB" : "#fff",
                    borderBottom: "1px solid #EAEEF2",
                    borderTop: rowIdx === lastIdx ? "3px solid #1863DC" : "none",
                  }}
                >
                  {isEditing && (
                    <td style={{ padding: "12px 8px", textAlign: "center", width: "40px" }}>
                      {data.rows.length > 1 && (
                        <button
                          onClick={() => deleteRow(slide.id, rowIdx)}
                          style={{
                            background: "transparent",
                            border: "none",
                            cursor: "pointer",
                            fontSize: "16px",
                            padding: "0",
                            color: "#DC2143",
                          }}
                          title="Delete row"
                        >
                          🗑️
                        </button>
                      )}
                    </td>
                  )}
                  {data.columns.map((col, colIdx) => {
                    const currentValue = row[col.key];
                    const isLatestRow = rowIdx === lastIdx;
                    const isFirstCol = colIdx === 0;
                    const isLastCol = colIdx === data.columns.length - 1;

                    return (
                      <td
                        key={col.key}
                        style={{
                          padding: "12px 16px",
                          fontWeight: isLatestRow ? "600" : "400",
                          color: "#212121",
                          borderLeft: isFirstCol && isLatestRow ? "3px solid #1863DC" : "none",
                          borderRight: isLastCol && isLatestRow ? "3px solid #1863DC" : "none",
                        }}
                      >
                        {isEditing ? (
                          <input
                            type="text"
                            defaultValue={currentValue}
                            onBlur={(e) => updateCellValue(slide.id, rowIdx, col.key, e.target.value)}
                            style={{ width: "100%", padding: "4px", border: "1px solid #DBDFE4", borderRadius: "4px" }}
                          />
                        ) : col.key === "sprint" ? (
                          `Sprint ${currentValue}`
                        ) : (
                          currentValue
                        )}
                      </td>
                    );
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    );
  };

  const Slide = ({ slide }) => (
    <div
      className="slide-section"
      style={{
        backgroundColor: "#fff",
        padding: "32px",
        borderRadius: "8px",
        boxShadow: "0 1px 3px rgba(0,0,0,0.08)",
        marginBottom: "24px",
        pageBreakAfter: "always",
      }}
    >
      {currentSprint > 0 && (
        <div
          style={{
            display: "inline-flex",
            alignItems: "center",
            gap: "8px",
            padding: "8px 16px",
            backgroundColor: "#EBF3FD",
            borderRadius: "6px",
            border: "2px solid #1863DC",
            marginBottom: "16px",
          }}
        >
          <span style={{ fontSize: "14px", fontWeight: "600", color: "#5A6872" }}>Sprint:</span>
          <span style={{ fontSize: "18px", fontWeight: "700", color: "#1863DC" }}>{currentSprint}</span>
        </div>
      )}
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          marginBottom: "20px",
          borderBottom: "2px solid #EAEEF2",
          paddingBottom: "12px",
        }}
      >
        <h2
          style={{
            fontSize: "20px",
            fontWeight: "600",
            color: "#212121",
            margin: 0,
            fontFamily: "Inter, -apple-system, BlinkMacSystemFont, sans-serif",
            flex: 1,
          }}
        >
          {isEditMode ? (
            <input
              key={`title-${slide.id}`}
              defaultValue={slide.title}
              onBlur={(e) => updateSlideTitle(slide.id, e.target.value)}
              style={{
                width: "100%",
                fontSize: "20px",
                fontWeight: "600",
                padding: "8px",
                border: "2px solid #1863DC",
                borderRadius: "4px",
              }}
            />
          ) : (
            slide.title
          )}
        </h2>
        <div style={{ display: "flex", gap: "8px", alignItems: "center" }}>
          {!isEditMode && slide.moreDetailsUrl && (
            <a
              href={slide.moreDetailsUrl}
              target="_blank"
              rel="noopener noreferrer"
              style={{
                padding: "8px 16px",
                fontSize: "13px",
                fontWeight: "600",
                backgroundColor: "#EBF3FD",
                color: "#1863DC",
                textDecoration: "none",
                borderRadius: "20px",
                whiteSpace: "nowrap",
                border: "1px solid #1863DC",
              }}
            >
              📊 More Details
            </a>
          )}
          {isEditMode && (
            <>
              <button
                onClick={() => openSlideImport(slide.id)}
                style={{
                  padding: "8px 16px",
                  fontSize: "13px",
                  fontWeight: "600",
                  backgroundColor: "#2DAD70",
                  color: "#fff",
                  border: "none",
                  borderRadius: "20px",
                  cursor: "pointer",
                  whiteSpace: "nowrap",
                }}
              >
                🔄 Update Slide
              </button>
              <input
                key={`url-${slide.id}`}
                type="url"
                placeholder="Google Sheet URL (required for updates)"
                defaultValue={slide.moreDetailsUrl || ""}
                onBlur={(e) => updateMoreDetailsUrl(slide.id, e.target.value)}
                style={{
                  width: "280px",
                  padding: "8px",
                  fontSize: "13px",
                  border: "1px solid #DBDFE4",
                  borderRadius: "4px",
                }}
              />
            </>
          )}
        </div>
      </div>
      {renderSlideContent(slide)}
    </div>
  );

  const renderSlideContent = (slide) => {
    switch (slide.type) {
      case "comparison":
        return (
          <div>
            {/* Bar Chart */}
            <div style={{ marginBottom: "32px" }}>
              {isEditMode && (
                <div style={{ display: "flex", justifyContent: "center", gap: "12px", marginBottom: "16px" }}>
                  <button
                    onClick={() => {
                      if (slide.data.sprints.length >= 5) {
                        alert("Maximum 5 sprints allowed");
                        return;
                      }
                      const newSlides = sprintData.slides.map((s) => {
                        if (s.id === slide.id && s.type === "comparison") {
                          return {
                            ...s,
                            data: {
                              ...s.data,
                              sprints: [...s.data.sprints, { sprintNumber: 0, paidUsers: 0, totalPaidQTD: 0 }],
                            },
                          };
                        }
                        return s;
                      }) as any;
                      preserveScroll();
                      setSprintData({ slides: newSlides });
                    }}
                    style={{
                      padding: "8px 16px",
                      backgroundColor: "#1863DC",
                      color: "white",
                      border: "none",
                      borderRadius: "4px",
                      cursor: "pointer",
                      fontSize: "13px",
                      fontWeight: "600",
                    }}
                  >
                    + Add Sprint ({slide.data.sprints.length}/5)
                  </button>
                  {slide.data.sprints.length > 1 && (
                    <button
                      onClick={() => {
                        const newSlides = sprintData.slides.map((s) => {
                          if (s.id === slide.id && s.type === "comparison") {
                            return {
                              ...s,
                              data: {
                                ...s.data,
                                sprints: s.data.sprints.slice(1), // Remove first (oldest) sprint
                              },
                            };
                          }
                          return s;
                        }) as any;
                        preserveScroll();
                        setSprintData({ slides: newSlides });
                      }}
                      style={{
                        padding: "8px 16px",
                        backgroundColor: "#DC2143",
                        color: "white",
                        border: "none",
                        borderRadius: "4px",
                        cursor: "pointer",
                        fontSize: "13px",
                        fontWeight: "600",
                      }}
                    >
                      − Remove Oldest Sprint
                    </button>
                  )}
                </div>
              )}
              <div
                style={{
                  display: "flex",
                  gap: "20px",
                  padding: "40px 40px 40px 80px",
                  backgroundColor: "#FFFFFF",
                  borderRadius: "8px",
                  position: "relative",
                }}
              >
                {/* Y-Axis */}
                <div
                  style={{
                    position: "absolute",
                    left: "10px",
                    top: "40px",
                    bottom: "80px",
                    display: "flex",
                    flexDirection: "column",
                    justifyContent: "space-between",
                    fontSize: "14px",
                    color: "#5A6872",
                  }}
                >
                  {[1400, 1050, 700, 350, 0].map((value) => (
                    <div key={value} style={{ textAlign: "right", width: "50px" }}>
                      {value}
                    </div>
                  ))}
                </div>

                {/* Grid Lines */}
                <div
                  style={{
                    position: "absolute",
                    left: "70px",
                    right: "40px",
                    top: "40px",
                    bottom: "80px",
                    display: "flex",
                    flexDirection: "column",
                    justifyContent: "space-between",
                  }}
                >
                  {[0, 1, 2, 3, 4].map((i) => (
                    <div key={i} style={{ borderTop: "1px solid #EAEEF2", width: "100%" }} />
                  ))}
                </div>

                {/* Chart Content */}
                <div
                  style={{
                    display: "flex",
                    gap: `${Math.max(120 - (slide.data.sprints.length - 2) * 20, 40)}px`,
                    alignItems: "flex-end",
                    justifyContent: "center",
                    flex: 1,
                    paddingBottom: "60px",
                    minHeight: "350px",
                    position: "relative",
                  }}
                >
                  {slide.data.sprints.map((sprint, idx) => {
                    const barWidth = Math.max(70 - (slide.data.sprints.length - 2) * 8, 40);
                    const barGap = Math.max(16 - (slide.data.sprints.length - 2) * 2, 8);

                    return (
                      <div
                        key={idx}
                        style={{ display: "flex", flexDirection: "column", alignItems: "center", position: "relative" }}
                      >
                        {/* Bars Container */}
                        <div style={{ display: "flex", gap: `${barGap}px`, alignItems: "flex-end" }}>
                          {/* Blue Bar - Paid users (this sprint) */}
                          <div style={{ display: "flex", flexDirection: "column", alignItems: "center" }}>
                            <div style={{ fontSize: "13px", fontWeight: "600", color: "#4A90E2", marginBottom: "8px" }}>
                              {isEditMode ? (
                                <input
                                  type="number"
                                  defaultValue={sprint.paidUsers}
                                  onBlur={(e) => {
                                    const newSprints = [...slide.data.sprints];
                                    newSprints[idx].paidUsers = parseInt(e.target.value) || 0;
                                    updateSlideData(slide.id, ["sprints"], newSprints);
                                  }}
                                  style={{
                                    width: "60px",
                                    padding: "4px",
                                    fontSize: "13px",
                                    border: "1px solid #4A90E2",
                                    borderRadius: "4px",
                                    textAlign: "center",
                                  }}
                                />
                              ) : (
                                sprint.paidUsers || ""
                              )}
                            </div>
                            <div
                              style={{
                                width: `${barWidth}px`,
                                height: `${Math.max((sprint.paidUsers / 1400) * 280, sprint.paidUsers > 0 ? 20 : 0)}px`,
                                backgroundColor: "#4A90E2",
                                borderRadius: "4px 4px 0 0",
                              }}
                            />
                          </div>

                          {/* Green Bar - Total paid users QTD */}
                          <div style={{ display: "flex", flexDirection: "column", alignItems: "center" }}>
                            <div style={{ fontSize: "13px", fontWeight: "600", color: "#5CB85C", marginBottom: "8px" }}>
                              {isEditMode ? (
                                <input
                                  type="number"
                                  defaultValue={sprint.totalPaidQTD}
                                  onBlur={(e) => {
                                    const newSprints = [...slide.data.sprints];
                                    newSprints[idx].totalPaidQTD = parseInt(e.target.value) || 0;
                                    updateSlideData(slide.id, ["sprints"], newSprints);
                                  }}
                                  style={{
                                    width: "60px",
                                    padding: "4px",
                                    fontSize: "13px",
                                    border: "1px solid #5CB85C",
                                    borderRadius: "4px",
                                    textAlign: "center",
                                  }}
                                />
                              ) : (
                                sprint.totalPaidQTD || ""
                              )}
                            </div>
                            <div
                              style={{
                                width: `${barWidth}px`,
                                height: `${Math.max((sprint.totalPaidQTD / 1400) * 280, sprint.totalPaidQTD > 0 ? 20 : 0)}px`,
                                backgroundColor: "#5CB85C",
                                borderRadius: "4px 4px 0 0",
                              }}
                            />
                          </div>
                        </div>

                        {/* Sprint Label at bottom - relative to this pair */}
                        <div
                          style={{
                            fontSize: "15px",
                            fontWeight: "600",
                            color: "#5A6872",
                            marginTop: "16px",
                            textAlign: "center",
                          }}
                        >
                          {isEditMode ? (
                            <input
                              type="text"
                              placeholder="Sprint #"
                              defaultValue={sprint.sprintNumber || ""}
                              onBlur={(e) => {
                                const newSprints = [...slide.data.sprints];
                                newSprints[idx].sprintNumber = parseInt(e.target.value) || 0;
                                updateSlideData(slide.id, ["sprints"], newSprints);
                              }}
                              style={{
                                width: "100px",
                                padding: "4px 8px",
                                fontSize: "15px",
                                border: "1px solid #DBDFE4",
                                borderRadius: "4px",
                                textAlign: "center",
                              }}
                            />
                          ) : sprint.sprintNumber ? (
                            `Sprint ${sprint.sprintNumber}`
                          ) : (
                            ""
                          )}
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>

              {/* Legend */}
              <div
                style={{
                  display: "flex",
                  justifyContent: "center",
                  gap: "32px",
                  marginTop: "20px",
                  fontSize: "13px",
                }}
              >
                <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
                  <div style={{ width: "14px", height: "14px", backgroundColor: "#4A90E2", borderRadius: "2px" }} />
                  <span style={{ color: "#5A6872" }}>Paid users (this sprint)</span>
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
                  <div style={{ width: "14px", height: "14px", backgroundColor: "#5CB85C", borderRadius: "2px" }} />
                  <span style={{ color: "#5A6872" }}>Total paid users QTD</span>
                </div>
              </div>
            </div>

            {/* Quarter Stats */}
            <div>
              {isEditMode && (
                <div style={{ display: "flex", justifyContent: "center", gap: "12px", marginBottom: "16px" }}>
                  <button
                    onClick={() => {
                      const newSlides = sprintData.slides.map((s) => {
                        if (s.id === slide.id && s.type === "comparison") {
                          return {
                            ...s,
                            data: {
                              ...s.data,
                              quarters: [
                                ...s.data.quarters,
                                { quarter: `Q${s.data.quarters.length + 1}`, total: 0, average: 0 },
                              ],
                            },
                          };
                        }
                        return s;
                      }) as any;
                      preserveScroll();
                      setSprintData({ slides: newSlides });
                    }}
                    style={{
                      padding: "8px 16px",
                      backgroundColor: "#2DAD70",
                      color: "white",
                      border: "none",
                      borderRadius: "4px",
                      cursor: "pointer",
                      fontSize: "13px",
                      fontWeight: "600",
                    }}
                  >
                    + Add Quarter
                  </button>
                  {slide.data.quarters.length > 1 && (
                    <button
                      onClick={() => {
                        const newSlides = sprintData.slides.map((s) => {
                          if (s.id === slide.id && s.type === "comparison") {
                            return {
                              ...s,
                              data: {
                                ...s.data,
                                quarters: s.data.quarters.slice(0, -1),
                              },
                            };
                          }
                          return s;
                        }) as any;
                        preserveScroll();
                        setSprintData({ slides: newSlides });
                      }}
                      style={{
                        padding: "8px 16px",
                        backgroundColor: "#DC2143",
                        color: "white",
                        border: "none",
                        borderRadius: "4px",
                        cursor: "pointer",
                        fontSize: "13px",
                        fontWeight: "600",
                      }}
                    >
                      − Remove Last Quarter
                    </button>
                  )}
                </div>
              )}
              <div style={{ display: "flex", gap: "24px", justifyContent: "center" }}>
                {slide.data.quarters.map((quarter, idx) => (
                  <div
                    key={idx}
                    style={{
                      padding: "24px 32px",
                      backgroundColor: "#FFFFFF",
                      border: "2px solid #1863DC",
                      borderRadius: "8px",
                      minWidth: "200px",
                    }}
                  >
                    <div
                      style={{
                        fontSize: "18px",
                        fontWeight: "600",
                        color: "#1863DC",
                        marginBottom: "16px",
                        textAlign: "center",
                      }}
                    >
                      {isEditMode ? (
                        <input
                          type="text"
                          defaultValue={quarter.quarter}
                          onBlur={(e) => {
                            const newQuarters = [...slide.data.quarters];
                            newQuarters[idx].quarter = e.target.value;
                            updateSlideData(slide.id, ["quarters"], newQuarters);
                          }}
                          style={{
                            width: "80px",
                            padding: "4px 8px",
                            fontSize: "18px",
                            border: "1px solid #1863DC",
                            borderRadius: "4px",
                            textAlign: "center",
                            fontWeight: "600",
                          }}
                        />
                      ) : (
                        quarter.quarter
                      )}
                    </div>

                    <div style={{ marginBottom: "12px" }}>
                      <div style={{ fontSize: "13px", color: "#5A6872", marginBottom: "4px" }}>Total:</div>
                      <div style={{ fontSize: "24px", fontWeight: "700", color: "#212121" }}>
                        {isEditMode ? (
                          <input
                            type="number"
                            defaultValue={quarter.total}
                            onBlur={(e) => {
                              const newQuarters = [...slide.data.quarters];
                              newQuarters[idx].total = parseInt(e.target.value) || 0;
                              updateSlideData(slide.id, ["quarters"], newQuarters);
                            }}
                            style={{
                              width: "120px",
                              padding: "4px 8px",
                              fontSize: "24px",
                              border: "1px solid #DBDFE4",
                              borderRadius: "4px",
                              fontWeight: "700",
                            }}
                          />
                        ) : (
                          quarter.total.toLocaleString()
                        )}
                      </div>
                    </div>

                    <div>
                      <div style={{ fontSize: "13px", color: "#5A6872", marginBottom: "4px" }}>Average:</div>
                      <div style={{ fontSize: "24px", fontWeight: "700", color: "#212121" }}>
                        {isEditMode ? (
                          <input
                            type="number"
                            defaultValue={quarter.average}
                            onBlur={(e) => {
                              const newQuarters = [...slide.data.quarters];
                              newQuarters[idx].average = parseInt(e.target.value) || 0;
                              updateSlideData(slide.id, ["quarters"], newQuarters);
                            }}
                            style={{
                              width: "120px",
                              padding: "4px 8px",
                              fontSize: "24px",
                              border: "1px solid #DBDFE4",
                              borderRadius: "4px",
                              fontWeight: "700",
                            }}
                          />
                        ) : (
                          quarter.average.toLocaleString()
                        )}
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          </div>
        );
      case "rankings":
        return (
          <div>
            <div style={{ marginBottom: "24px" }}>
              <h3
                style={{
                  fontSize: "14px",
                  fontWeight: "600",
                  marginBottom: "12px",
                  color: "#5A6872",
                  textTransform: "uppercase",
                  letterSpacing: "0.5px",
                }}
              >
                TOTAL #1 RANKINGS
              </h3>
              <div style={{ fontSize: "32px", fontWeight: "700", color: "#1863DC" }}>
                {isEditMode ? (
                  <input
                    key={`total-${slide.id}`}
                    type="number"
                    defaultValue={slide.data.total}
                    onBlur={(e) => updateSlideData(slide.id, ["total"], e.target.value)}
                    style={{
                      fontSize: "32px",
                      padding: "8px",
                      width: "120px",
                      border: "2px solid #1863DC",
                      borderRadius: "4px",
                    }}
                  />
                ) : (
                  slide.data.total
                )}
              </div>
            </div>
            <div style={{ marginBottom: "24px" }}>
              <div
                style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "12px" }}
              >
                <h3
                  style={{
                    fontSize: "14px",
                    fontWeight: "600",
                    color: "#5A6872",
                    textTransform: "uppercase",
                    letterSpacing: "0.5px",
                  }}
                >
                  #1 Rankings by Region
                </h3>
                {isEditMode && (
                  <button
                    onClick={() => {
                      const newSlides = sprintData.slides.map((s) => {
                        if (s.id === slide.id && s.type === "rankings") {
                          return {
                            ...s,
                            data: {
                              ...s.data,
                              byRegion: [...s.data.byRegion, { region: "New Region", count: 0, keywords: "" }],
                            },
                          };
                        }
                        return s;
                      }) as any;
                      preserveScroll();
                      setSprintData({ slides: newSlides });
                    }}
                    style={{
                      padding: "6px 12px",
                      backgroundColor: "#1863DC",
                      color: "white",
                      border: "none",
                      borderRadius: "4px",
                      cursor: "pointer",
                      fontSize: "12px",
                    }}
                  >
                    + Add Region
                  </button>
                )}
              </div>
              {slide.data.byRegion.map((item, idx) => (
                <div
                  key={idx}
                  style={{
                    marginBottom: "12px",
                    padding: "12px 16px",
                    backgroundColor: "#F8FAFB",
                    borderRadius: "6px",
                    border: "1px solid #EAEEF2",
                    position: "relative",
                  }}
                >
                  {isEditMode ? (
                    <div style={{ display: "flex", flexDirection: "column", gap: "8px" }}>
                      <div style={{ display: "flex", gap: "8px", alignItems: "center" }}>
                        <input
                          type="text"
                          defaultValue={item.region}
                          onBlur={(e) => {
                            const newRegion = [...slide.data.byRegion];
                            newRegion[idx].region = e.target.value;
                            updateSlideData(slide.id, ["byRegion"], newRegion);
                          }}
                          placeholder="Region"
                          style={{
                            padding: "6px",
                            fontSize: "14px",
                            border: "1px solid #DBDFE4",
                            borderRadius: "4px",
                            fontWeight: "600",
                            flex: "1",
                          }}
                        />
                        <input
                          type="number"
                          defaultValue={item.count}
                          onBlur={(e) => {
                            const newRegion = [...slide.data.byRegion];
                            newRegion[idx].count = parseInt(e.target.value) || 0;
                            updateSlideData(slide.id, ["byRegion"], newRegion);
                          }}
                          placeholder="Count"
                          style={{
                            padding: "6px",
                            fontSize: "14px",
                            border: "1px solid #DBDFE4",
                            borderRadius: "4px",
                            width: "80px",
                          }}
                        />
                        <button
                          onClick={() => {
                            const newRegion = slide.data.byRegion.filter((_, i) => i !== idx);
                            updateSlideData(slide.id, ["byRegion"], newRegion);
                          }}
                          style={{
                            padding: "6px 10px",
                            backgroundColor: "#DC2143",
                            color: "white",
                            border: "none",
                            borderRadius: "4px",
                            cursor: "pointer",
                            fontSize: "12px",
                          }}
                        >
                          ✕
                        </button>
                      </div>
                      <textarea
                        defaultValue={item.keywords}
                        onBlur={(e) => {
                          const newRegion = [...slide.data.byRegion];
                          newRegion[idx].keywords = e.target.value;
                          updateSlideData(slide.id, ["byRegion"], newRegion);
                        }}
                        placeholder="Keywords (comma separated)"
                        style={{
                          padding: "6px",
                          fontSize: "12px",
                          border: "1px solid #DBDFE4",
                          borderRadius: "4px",
                          minHeight: "60px",
                          resize: "vertical",
                        }}
                      />
                    </div>
                  ) : (
                    <>
                      <strong style={{ color: "#212121", fontSize: "14px" }}>
                        {item.region} ({item.count})
                      </strong>
                      <div style={{ fontSize: "12px", color: "#5A6872", marginTop: "8px", lineHeight: "1.8" }}>
                        {item.keywords.split(",").map((keyword, i) => (
                          <div key={i}>{keyword.trim()}</div>
                        ))}
                      </div>
                    </>
                  )}
                </div>
              ))}
            </div>
            <div style={{ marginBottom: "24px" }}>
              <h3
                style={{
                  fontSize: "14px",
                  fontWeight: "600",
                  marginBottom: "12px",
                  color: "#5A6872",
                  textTransform: "uppercase",
                  letterSpacing: "0.5px",
                }}
              >
                Position Changes
              </h3>
              <EditableTable slide={{ ...slide, id: slide.id, data: slide.data.positionChanges }} />
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "20px" }}>
              <div>
                <div
                  style={{
                    display: "flex",
                    justifyContent: "space-between",
                    alignItems: "center",
                    marginBottom: "12px",
                  }}
                >
                  <h3 style={{ fontSize: "14px", fontWeight: "600", color: "#2DAD70" }}>↑ Improved</h3>
                  {isEditMode && (
                    <button
                      onClick={() => {
                        const newSlides = sprintData.slides.map((s) => {
                          if (s.id === slide.id && s.type === "rankings") {
                            return {
                              ...s,
                              data: {
                                ...s.data,
                                improved: [...s.data.improved, { region: "New Region", count: 0, keywords: "" }],
                              },
                            };
                          }
                          return s;
                        }) as any;
                        preserveScroll();
                        setSprintData({ slides: newSlides });
                      }}
                      style={{
                        padding: "4px 8px",
                        backgroundColor: "#2DAD70",
                        color: "white",
                        border: "none",
                        borderRadius: "4px",
                        cursor: "pointer",
                        fontSize: "11px",
                      }}
                    >
                      + Add
                    </button>
                  )}
                </div>
                {slide.data.improved.map((item, idx) => (
                  <div
                    key={idx}
                    style={{
                      marginBottom: "12px",
                      padding: "12px 16px",
                      backgroundColor: "#E6F7ED",
                      borderRadius: "6px",
                      border: "1px solid #2DAD70",
                    }}
                  >
                    {isEditMode ? (
                      <div style={{ display: "flex", flexDirection: "column", gap: "8px" }}>
                        <div style={{ display: "flex", gap: "8px", alignItems: "center" }}>
                          <input
                            type="text"
                            defaultValue={item.region}
                            onBlur={(e) => {
                              const newImproved = [...slide.data.improved];
                              newImproved[idx].region = e.target.value;
                              updateSlideData(slide.id, ["improved"], newImproved);
                            }}
                            placeholder="Region"
                            style={{
                              padding: "6px",
                              fontSize: "12px",
                              border: "1px solid #DBDFE4",
                              borderRadius: "4px",
                              flex: "1",
                            }}
                          />
                          <input
                            type="number"
                            defaultValue={item.count}
                            onBlur={(e) => {
                              const newImproved = [...slide.data.improved];
                              newImproved[idx].count = parseInt(e.target.value) || 0;
                              updateSlideData(slide.id, ["improved"], newImproved);
                            }}
                            placeholder="Count"
                            style={{
                              padding: "6px",
                              fontSize: "12px",
                              border: "1px solid #DBDFE4",
                              borderRadius: "4px",
                              width: "60px",
                            }}
                          />
                          <button
                            onClick={() => {
                              const newImproved = slide.data.improved.filter((_, i) => i !== idx);
                              updateSlideData(slide.id, ["improved"], newImproved);
                            }}
                            style={{
                              padding: "4px 8px",
                              backgroundColor: "#DC2143",
                              color: "white",
                              border: "none",
                              borderRadius: "4px",
                              cursor: "pointer",
                              fontSize: "11px",
                            }}
                          >
                            ✕
                          </button>
                        </div>
                        <textarea
                          defaultValue={item.keywords}
                          onBlur={(e) => {
                            const newImproved = [...slide.data.improved];
                            newImproved[idx].keywords = e.target.value;
                            updateSlideData(slide.id, ["improved"], newImproved);
                          }}
                          placeholder="Keywords"
                          style={{
                            padding: "6px",
                            fontSize: "11px",
                            border: "1px solid #DBDFE4",
                            borderRadius: "4px",
                            minHeight: "50px",
                            resize: "vertical",
                          }}
                        />
                      </div>
                    ) : (
                      <>
                        <strong style={{ color: "#212121", fontSize: "14px" }}>
                          {item.region} ({item.count})
                        </strong>
                        <div style={{ fontSize: "12px", color: "#5A6872", marginTop: "8px", lineHeight: "1.8" }}>
                          {item.keywords.split(",").map((keyword, i) => (
                            <div key={i}>{keyword.trim()}</div>
                          ))}
                        </div>
                      </>
                    )}
                  </div>
                ))}
              </div>
              <div>
                <div
                  style={{
                    display: "flex",
                    justifyContent: "space-between",
                    alignItems: "center",
                    marginBottom: "12px",
                  }}
                >
                  <h3 style={{ fontSize: "14px", fontWeight: "600", color: "#DC2143" }}>↓ Declined</h3>
                  {isEditMode && (
                    <button
                      onClick={() => {
                        const newSlides = sprintData.slides.map((s) => {
                          if (s.id === slide.id && s.type === "rankings") {
                            return {
                              ...s,
                              data: {
                                ...s.data,
                                declined: [...s.data.declined, { region: "New Region", count: 0, keywords: "" }],
                              },
                            };
                          }
                          return s;
                        }) as any;
                        preserveScroll();
                        setSprintData({ slides: newSlides });
                      }}
                      style={{
                        padding: "4px 8px",
                        backgroundColor: "#DC2143",
                        color: "white",
                        border: "none",
                        borderRadius: "4px",
                        cursor: "pointer",
                        fontSize: "11px",
                      }}
                    >
                      + Add
                    </button>
                  )}
                </div>
                {slide.data.declined.map((item, idx) => (
                  <div
                    key={idx}
                    style={{
                      marginBottom: "12px",
                      padding: "12px 16px",
                      backgroundColor: "#FEE9E9",
                      borderRadius: "6px",
                      border: "1px solid #DC2143",
                    }}
                  >
                    {isEditMode ? (
                      <div style={{ display: "flex", flexDirection: "column", gap: "8px" }}>
                        <div style={{ display: "flex", gap: "8px", alignItems: "center" }}>
                          <input
                            type="text"
                            defaultValue={item.region}
                            onBlur={(e) => {
                              const newDeclined = [...slide.data.declined];
                              newDeclined[idx].region = e.target.value;
                              updateSlideData(slide.id, ["declined"], newDeclined);
                            }}
                            placeholder="Region"
                            style={{
                              padding: "6px",
                              fontSize: "12px",
                              border: "1px solid #DBDFE4",
                              borderRadius: "4px",
                              flex: "1",
                            }}
                          />
                          <input
                            type="number"
                            defaultValue={item.count}
                            onBlur={(e) => {
                              const newDeclined = [...slide.data.declined];
                              newDeclined[idx].count = parseInt(e.target.value) || 0;
                              updateSlideData(slide.id, ["declined"], newDeclined);
                            }}
                            placeholder="Count"
                            style={{
                              padding: "6px",
                              fontSize: "12px",
                              border: "1px solid #DBDFE4",
                              borderRadius: "4px",
                              width: "60px",
                            }}
                          />
                          <button
                            onClick={() => {
                              const newDeclined = slide.data.declined.filter((_, i) => i !== idx);
                              updateSlideData(slide.id, ["declined"], newDeclined);
                            }}
                            style={{
                              padding: "4px 8px",
                              backgroundColor: "#DC2143",
                              color: "white",
                              border: "none",
                              borderRadius: "4px",
                              cursor: "pointer",
                              fontSize: "11px",
                            }}
                          >
                            ✕
                          </button>
                        </div>
                        <textarea
                          defaultValue={item.keywords}
                          onBlur={(e) => {
                            const newDeclined = [...slide.data.declined];
                            newDeclined[idx].keywords = e.target.value;
                            updateSlideData(slide.id, ["declined"], newDeclined);
                          }}
                          placeholder="Keywords"
                          style={{
                            padding: "6px",
                            fontSize: "11px",
                            border: "1px solid #DBDFE4",
                            borderRadius: "4px",
                            minHeight: "50px",
                            resize: "vertical",
                          }}
                        />
                      </div>
                    ) : (
                      <>
                        <strong style={{ color: "#212121", fontSize: "14px" }}>
                          {item.region} ({item.count})
                        </strong>
                        <div style={{ fontSize: "12px", color: "#5A6872", marginTop: "8px", lineHeight: "1.8" }}>
                          {item.keywords.split(",").map((keyword, i) => (
                            <div key={i}>{keyword.trim()}</div>
                          ))}
                        </div>
                      </>
                    )}
                  </div>
                ))}
              </div>
            </div>
          </div>
        );

      case "table":
        return <EditableTable slide={slide} />;

      case "pluginRanking":
        return (
          <div>
            <div
              style={{
                marginBottom: "20px",
                padding: "16px",
                backgroundColor: "#EBF3FD",
                borderRadius: "6px",
                border: "1px solid #1863DC",
              }}
            >
              <strong style={{ color: "#212121", fontSize: "14px" }}>QUARTER TARGET: </strong>
              {isEditMode ? (
                <input
                  type="number"
                  value={slide.data.quarterTarget}
                  onChange={(e) => updateSlideData(slide.id, ["quarterTarget"], e.target.value)}
                  style={{ padding: "6px 8px", width: "80px", border: "1px solid #1863DC", borderRadius: "4px" }}
                />
              ) : (
                <span style={{ fontWeight: "600", color: "#1863DC", fontSize: "16px" }}>
                  {slide.data.quarterTarget}
                </span>
              )}
            </div>
            <EditableTable slide={slide} />
          </div>
        );

      case "supportData":
        return (
          <div>
            <div style={{ marginBottom: "32px" }}>
              <h3 style={{ fontSize: "16px", fontWeight: "600", marginBottom: "16px", color: "#212121" }}>Tickets</h3>
              <EditableTable slide={{ ...slide, id: slide.id, data: slide.data.tickets, type: slide.type }} />
            </div>
            <div>
              <h3 style={{ fontSize: "16px", fontWeight: "600", marginBottom: "16px", color: "#212121" }}>Live Chat</h3>
              <EditableTable slide={{ ...slide, id: slide.id + 1000, data: slide.data.liveChat, type: slide.type }} />
            </div>
          </div>
        );

      case "agencyLeads":
        return (
          <div>
            <div style={{ marginBottom: "32px" }}>
              <EditableTable slide={{ ...slide, id: slide.id, data: slide.data.leadsConversion, type: slide.type }} />
            </div>
            <div>
              <EditableTable
                slide={{ ...slide, id: slide.id + 2000, data: slide.data.q3Performance, type: slide.type }}
              />
            </div>
          </div>
        );

      case "googleAdsObservations":
        return (
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "24px" }}>
            {/* Performance Section */}
            <div
              style={{
                padding: "24px",
                backgroundColor: "#FFFFFF",
                borderRadius: "8px",
                border: "2px solid #1863DC",
              }}
            >
              <h3
                style={{
                  fontSize: "16px",
                  fontWeight: "600",
                  color: "#1863DC",
                  marginBottom: "20px",
                  textTransform: "uppercase",
                  letterSpacing: "0.5px",
                }}
              >
                Performance (Current Sprint vs. Last Sprint)
              </h3>
              <div style={{ display: "flex", flexDirection: "column", gap: "16px" }}>
                {["impressions", "clicks", "CTR", "conversions"].map((metric) => {
                  const metricData = slide.data.performance[metric];
                  const label = metric.charAt(0).toUpperCase() + metric.slice(1);
                  return (
                    <div
                      key={metric}
                      style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}
                    >
                      <span style={{ fontSize: "15px", fontWeight: "600", color: "#212121" }}>{label}</span>
                      <div style={{ display: "flex", alignItems: "center", gap: "12px" }}>
                        {isEditMode ? (
                          <>
                            <input
                              type="number"
                              step="0.1"
                              defaultValue={metricData.value}
                              onBlur={(e) =>
                                updateSlideData(
                                  slide.id,
                                  ["performance", metric, "value"],
                                  parseFloat(e.target.value) || 0,
                                )
                              }
                              style={{
                                width: "100px",
                                padding: "6px 8px",
                                border: "1px solid #DFE3E8",
                                borderRadius: "4px",
                                fontSize: "14px",
                              }}
                            />
                            <select
                              value={metricData.isIncrease ? "increase" : "decrease"}
                              onChange={(e) =>
                                updateSlideData(
                                  slide.id,
                                  ["performance", metric, "isIncrease"],
                                  e.target.value === "increase",
                                )
                              }
                              style={{
                                padding: "6px 8px",
                                border: "1px solid #DFE3E8",
                                borderRadius: "4px",
                                fontSize: "14px",
                                backgroundColor: "#fff",
                                cursor: "pointer",
                              }}
                            >
                              <option value="increase">↑ Increase</option>
                              <option value="decrease">↓ Decrease</option>
                            </select>
                          </>
                        ) : (
                          <span
                            style={{
                              fontSize: "15px",
                              fontWeight: "700",
                              color: metricData.isIncrease ? "#2DAD70" : "#E94B4B",
                              display: "flex",
                              alignItems: "center",
                              gap: "4px",
                            }}
                          >
                            {metricData.isIncrease ? "↑" : "↓"} {metricData.value}
                          </span>
                        )}
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>

            {/* Auction Insights Section */}
            <div
              style={{
                padding: "24px",
                backgroundColor: "#FFFFFF",
                borderRadius: "8px",
                border: "2px solid #1863DC",
              }}
            >
              <h3
                style={{
                  fontSize: "16px",
                  fontWeight: "600",
                  color: "#1863DC",
                  marginBottom: "20px",
                  textTransform: "uppercase",
                  letterSpacing: "0.5px",
                }}
              >
                Auction Insights
              </h3>
              <div style={{ display: "flex", flexDirection: "column", gap: "16px" }}>
                {[
                  { key: "impressionShare", label: "Impression Share" },
                  { key: "absoluteTopOfPage", label: "Absolute Top of Page" },
                ].map(({ key, label }) => {
                  const metricData = slide.data.auctionInsights[key];
                  return (
                    <div key={key} style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                      <span style={{ fontSize: "15px", fontWeight: "600", color: "#212121" }}>{label}</span>
                      <div style={{ display: "flex", alignItems: "center", gap: "12px" }}>
                        {isEditMode ? (
                          <>
                            <input
                              type="number"
                              step="0.1"
                              defaultValue={metricData.value}
                              onBlur={(e) =>
                                updateSlideData(
                                  slide.id,
                                  ["auctionInsights", key, "value"],
                                  parseFloat(e.target.value) || 0,
                                )
                              }
                              style={{
                                width: "100px",
                                padding: "6px 8px",
                                border: "1px solid #DFE3E8",
                                borderRadius: "4px",
                                fontSize: "14px",
                              }}
                            />
                            <select
                              value={metricData.isIncrease ? "increase" : "decrease"}
                              onChange={(e) =>
                                updateSlideData(
                                  slide.id,
                                  ["auctionInsights", key, "isIncrease"],
                                  e.target.value === "increase",
                                )
                              }
                              style={{
                                padding: "6px 8px",
                                border: "1px solid #DFE3E8",
                                borderRadius: "4px",
                                fontSize: "14px",
                                backgroundColor: "#fff",
                                cursor: "pointer",
                              }}
                            >
                              <option value="increase">↑ Increase</option>
                              <option value="decrease">↓ Decrease</option>
                            </select>
                          </>
                        ) : (
                          <span
                            style={{
                              fontSize: "15px",
                              fontWeight: "700",
                              color: metricData.isIncrease ? "#2DAD70" : "#E94B4B",
                              display: "flex",
                              alignItems: "center",
                              gap: "4px",
                            }}
                          >
                            {metricData.isIncrease ? "↑" : "↓"} {metricData.value}
                          </span>
                        )}
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          </div>
        );

      case "textarea":
        return (
          <div
            style={{
              padding: "20px",
              backgroundColor: "#E6F7ED",
              borderRadius: "8px",
              border: "1px solid #2DAD70",
              fontSize: "14px",
              whiteSpace: "pre-line",
              color: "#212121",
              lineHeight: "1.8",
              minHeight: "200px",
            }}
          >
            {isEditMode ? (
              <textarea
                key={`textarea-${slide.id}`}
                defaultValue={slide.data.text}
                onBlur={(e) => updateSlideData(slide.id, ["text"], e.target.value)}
                style={{
                  width: "100%",
                  height: "350px",
                  padding: "12px",
                  fontSize: "14px",
                  border: "1px solid #2DAD70",
                  borderRadius: "6px",
                  fontFamily: "Inter, sans-serif",
                }}
              />
            ) : (
              slide.data.text || "No content yet. Click Edit to add observations."
            )}
          </div>
        );

      case "quarterStats":
        return (
          <div>
            <div style={{ marginBottom: "20px" }}>
              <div
                style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "16px" }}
              >
                <h3 style={{ fontSize: "16px", fontWeight: "600", color: "#212121" }}>Quarter Stats</h3>
                {isEditMode && (
                  <div style={{ display: "flex", gap: "8px" }}>
                    {editingStatsId !== slide.id && (
                      <button
                        onClick={() => setEditingStatsId(slide.id)}
                        style={{
                          padding: "6px 14px",
                          fontSize: "12px",
                          fontWeight: "600",
                          backgroundColor: "#1863DC",
                          color: "#fff",
                          border: "none",
                          borderRadius: "4px",
                          cursor: "pointer",
                        }}
                      >
                        ✏️ Edit Stats
                      </button>
                    )}
                    {editingStatsId === slide.id && (
                      <button
                        onClick={() => setEditingStatsId(null)}
                        style={{
                          padding: "6px 14px",
                          fontSize: "12px",
                          fontWeight: "600",
                          backgroundColor: "#2DAD70",
                          color: "#fff",
                          border: "none",
                          borderRadius: "4px",
                          cursor: "pointer",
                        }}
                      >
                        ✓ Done
                      </button>
                    )}
                    <button
                      onClick={() => openSlideImport(slide.id)}
                      style={{
                        padding: "6px 14px",
                        fontSize: "12px",
                        fontWeight: "600",
                        backgroundColor: "#2DAD70",
                        color: "#fff",
                        border: "none",
                        borderRadius: "4px",
                        cursor: "pointer",
                      }}
                    >
                      🔄 Update Stats
                    </button>
                  </div>
                )}
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(5, 1fr)", gap: "16px" }}>
                {Object.entries(slide.data.quarterStats).map(([key, value]) => (
                  <div
                    key={key}
                    style={{
                      padding: "20px",
                      backgroundColor: "#EBF3FD",
                      borderRadius: "8px",
                      textAlign: "center",
                      border: "1px solid #4682E1",
                    }}
                  >
                    <div style={{ fontSize: "36px", fontWeight: "700", color: "#1863DC", marginBottom: "8px" }}>
                      {editingStatsId === slide.id ? (
                        <input
                          type="number"
                          defaultValue={Number(value)}
                          onBlur={(e) => updateSlideData(slide.id, ["quarterStats", key], e.target.value)}
                          style={{
                            width: "100%",
                            fontSize: "32px",
                            padding: "4px",
                            border: "2px solid #1863DC",
                            borderRadius: "4px",
                            textAlign: "center",
                          }}
                        />
                      ) : (
                        <span>{String(value)}</span>
                      )}
                    </div>
                    <div
                      style={{
                        fontSize: "11px",
                        color: "#5A6872",
                        textTransform: "uppercase",
                        fontWeight: "600",
                        letterSpacing: "0.5px",
                      }}
                    >
                      {key.replace(/([A-Z])/g, " $1").replace(/^./, (str) => str.toUpperCase())}
                    </div>
                  </div>
                ))}
              </div>
            </div>
            <EditableTable slide={slide} />
          </div>
        );

      case "withTarget":
        return (
          <div>
            <div
              style={{
                marginBottom: "20px",
                padding: "16px",
                backgroundColor: "#EBF3FD",
                borderRadius: "6px",
                border: "1px solid #1863DC",
              }}
            >
              <strong style={{ color: "#212121", fontSize: "14px" }}>
                {slide.data.targetLabel || "SPRINT TARGET"}:{" "}
              </strong>
              {isEditMode ? (
                <input
                  type="number"
                  value={slide.data.target}
                  onChange={(e) => updateSlideData(slide.id, ["target"], e.target.value)}
                  style={{ padding: "6px 8px", width: "80px", border: "1px solid #1863DC", borderRadius: "4px" }}
                />
              ) : (
                <span style={{ fontWeight: "600", color: "#1863DC", fontSize: "16px" }}>{slide.data.target}</span>
              )}
            </div>
            <EditableTable slide={slide} />
            {slide.data.total && !slide.data.hideQtdStats && (
              <div style={{ marginTop: "20px" }}>
                <div
                  style={{
                    display: "flex",
                    justifyContent: "space-between",
                    alignItems: "center",
                    marginBottom: "16px",
                  }}
                >
                  <h3 style={{ fontSize: "16px", fontWeight: "600", color: "#212121" }}>QTD Stats</h3>
                  {isEditMode && (
                    <div style={{ display: "flex", gap: "8px" }}>
                      {editingStatsId !== slide.id + 100 && (
                        <button
                          onClick={() => setEditingStatsId(slide.id + 100)}
                          style={{
                            padding: "6px 14px",
                            fontSize: "12px",
                            fontWeight: "600",
                            backgroundColor: "#1863DC",
                            color: "#fff",
                            border: "none",
                            borderRadius: "4px",
                            cursor: "pointer",
                          }}
                        >
                          ✏️ Edit Stats
                        </button>
                      )}
                      {editingStatsId === slide.id + 100 && (
                        <button
                          onClick={() => setEditingStatsId(null)}
                          style={{
                            padding: "6px 14px",
                            fontSize: "12px",
                            fontWeight: "600",
                            backgroundColor: "#2DAD70",
                            color: "#fff",
                            border: "none",
                            borderRadius: "4px",
                            cursor: "pointer",
                          }}
                        >
                          ✓ Done
                        </button>
                      )}
                      <button
                        onClick={() => openSlideImport(slide.id)}
                        style={{
                          padding: "6px 14px",
                          fontSize: "12px",
                          fontWeight: "600",
                          backgroundColor: "#2DAD70",
                          color: "#fff",
                          border: "none",
                          borderRadius: "4px",
                          cursor: "pointer",
                        }}
                      >
                        🔄 Update Stats
                      </button>
                    </div>
                  )}
                </div>
                <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: "16px" }}>
                  {Object.entries(slide.data.total).map(([key, value]) => (
                    <div
                      key={key}
                      style={{
                        padding: "20px",
                        backgroundColor: "#EBF3FD",
                        borderRadius: "8px",
                        textAlign: "center",
                        border: "1px solid #4682E1",
                      }}
                    >
                      <div style={{ fontSize: "36px", fontWeight: "700", color: "#1863DC", marginBottom: "8px" }}>
                        {editingStatsId === slide.id + 100 ? (
                          <input
                            type="number"
                            defaultValue={Number(value)}
                            onBlur={(e) => updateSlideData(slide.id, ["total", key], e.target.value)}
                            style={{
                              width: "100%",
                              fontSize: "32px",
                              padding: "4px",
                              border: "2px solid #1863DC",
                              borderRadius: "4px",
                              textAlign: "center",
                            }}
                          />
                        ) : (
                          <span>{String(value)}</span>
                        )}
                      </div>
                      <div
                        style={{
                          fontSize: "11px",
                          color: "#5A6872",
                          textTransform: "uppercase",
                          fontWeight: "600",
                          letterSpacing: "0.5px",
                        }}
                      >
                        {key}
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
        );

      case "referral":
        return (
          <div>
            {slide.data.target !== undefined && (
              <div
                style={{
                  marginBottom: "20px",
                  padding: "16px",
                  backgroundColor: "#EBF3FD",
                  borderRadius: "6px",
                  border: "1px solid #1863DC",
                }}
              >
                <strong style={{ color: "#212121", fontSize: "14px" }}>TARGET VS CURRENT - PAID PLANS: </strong>
                {isEditMode ? (
                  <>
                    <input
                      type="number"
                      value={slide.data.target}
                      onChange={(e) => updateSlideData(slide.id, ["target"], e.target.value)}
                      style={{
                        padding: "6px 8px",
                        width: "80px",
                        border: "1px solid #1863DC",
                        borderRadius: "4px",
                        marginRight: "8px",
                      }}
                    />
                    <span style={{ fontWeight: "600", color: "#1863DC", fontSize: "16px" }}>Vs</span>
                    <input
                      type="number"
                      value={slide.data.current}
                      onChange={(e) => updateSlideData(slide.id, ["current"], e.target.value)}
                      style={{
                        padding: "6px 8px",
                        width: "80px",
                        border: "1px solid #1863DC",
                        borderRadius: "4px",
                        marginLeft: "8px",
                      }}
                    />
                  </>
                ) : (
                  <span style={{ fontWeight: "600", color: "#1863DC", fontSize: "16px" }}>
                    {slide.data.target} Vs {slide.data.current}
                  </span>
                )}
              </div>
            )}
            <EditableTable slide={slide} />
            <div style={{ marginTop: "20px" }}>
              <div
                style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "16px" }}
              >
                <h3 style={{ fontSize: "16px", fontWeight: "600", color: "#212121" }}>App Lifetime Stats</h3>
                {isEditMode && (
                  <div style={{ display: "flex", gap: "8px" }}>
                    {editingStatsId !== slide.id + 200 && (
                      <button
                        onClick={() => setEditingStatsId(slide.id + 200)}
                        style={{
                          padding: "6px 14px",
                          fontSize: "12px",
                          fontWeight: "600",
                          backgroundColor: "#1863DC",
                          color: "#fff",
                          border: "none",
                          borderRadius: "4px",
                          cursor: "pointer",
                        }}
                      >
                        ✏️ Edit Stats
                      </button>
                    )}
                    {editingStatsId === slide.id + 200 && (
                      <button
                        onClick={() => setEditingStatsId(null)}
                        style={{
                          padding: "6px 14px",
                          fontSize: "12px",
                          fontWeight: "600",
                          backgroundColor: "#2DAD70",
                          color: "#fff",
                          border: "none",
                          borderRadius: "4px",
                          cursor: "pointer",
                        }}
                      >
                        ✓ Done
                      </button>
                    )}
                    <button
                      onClick={() => openSlideImport(slide.id)}
                      style={{
                        padding: "6px 14px",
                        fontSize: "12px",
                        fontWeight: "600",
                        backgroundColor: "#2DAD70",
                        color: "#fff",
                        border: "none",
                        borderRadius: "4px",
                        cursor: "pointer",
                      }}
                    >
                      🔄 Update Stats
                    </button>
                  </div>
                )}
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(2, 1fr)", gap: "16px" }}>
                {Object.entries(slide.data.lifetime).map(([key, value]) => (
                  <div
                    key={key}
                    style={{
                      padding: "20px",
                      backgroundColor: "#EBF3FD",
                      borderRadius: "8px",
                      textAlign: "center",
                      border: "1px solid #4682E1",
                    }}
                  >
                    <div style={{ fontSize: "36px", fontWeight: "700", color: "#1863DC", marginBottom: "8px" }}>
                      {editingStatsId === slide.id + 200 ? (
                        <input
                          type="number"
                          defaultValue={Number(value)}
                          onBlur={(e) => updateSlideData(slide.id, ["lifetime", key], e.target.value)}
                          style={{
                            width: "100%",
                            fontSize: "32px",
                            padding: "4px",
                            border: "2px solid #1863DC",
                            borderRadius: "4px",
                            textAlign: "center",
                          }}
                        />
                      ) : (
                        <span>{String(value)}</span>
                      )}
                    </div>
                    <div
                      style={{
                        fontSize: "11px",
                        color: "#5A6872",
                        textTransform: "uppercase",
                        fontWeight: "600",
                        letterSpacing: "0.5px",
                      }}
                    >
                      {key === "paid" ? "Paid Signups" : key === "trials" ? "Trial Signups" : key === "referrals" ? "Referrals" : key}
                    </div>
                  </div>
                ))}
              </div>
            </div>
          </div>
        );

      case "wixApp":
        return (
          <div>
            {slide.data.target !== undefined && (
              <div
                style={{
                  marginBottom: "20px",
                  padding: "16px",
                  backgroundColor: "#EBF3FD",
                  borderRadius: "6px",
                  border: "1px solid #1863DC",
                }}
              >
                <strong style={{ color: "#212121", fontSize: "14px" }}>TARGET VS CURRENT - PAID PLANS: </strong>
                {isEditMode ? (
                  <>
                    <input
                      type="number"
                      value={slide.data.target}
                      onChange={(e) => updateSlideData(slide.id, ["target"], e.target.value)}
                      style={{
                        padding: "6px 8px",
                        width: "80px",
                        border: "1px solid #1863DC",
                        borderRadius: "4px",
                        marginRight: "8px",
                      }}
                    />
                    <span style={{ fontWeight: "600", color: "#1863DC", fontSize: "16px" }}>Vs</span>
                    <input
                      type="number"
                      value={slide.data.current}
                      onChange={(e) => updateSlideData(slide.id, ["current"], e.target.value)}
                      style={{
                        padding: "6px 8px",
                        width: "80px",
                        border: "1px solid #1863DC",
                        borderRadius: "4px",
                        marginLeft: "8px",
                      }}
                    />
                  </>
                ) : (
                  <span style={{ fontWeight: "600", color: "#1863DC", fontSize: "16px" }}>
                    {slide.data.target} Vs {slide.data.current}
                  </span>
                )}
              </div>
            )}
            <EditableTable slide={slide} />
            <div style={{ marginTop: "20px" }}>
              <div
                style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "16px" }}
              >
                <h3 style={{ fontSize: "16px", fontWeight: "600", color: "#212121" }}>App Lifetime Stats</h3>
                {isEditMode && (
                  <div style={{ display: "flex", gap: "8px" }}>
                    {editingStatsId !== slide.id + 300 && (
                      <button
                        onClick={() => setEditingStatsId(slide.id + 300)}
                        style={{
                          padding: "6px 14px",
                          fontSize: "12px",
                          fontWeight: "600",
                          backgroundColor: "#1863DC",
                          color: "#fff",
                          border: "none",
                          borderRadius: "4px",
                          cursor: "pointer",
                        }}
                      >
                        ✏️ Edit Stats
                      </button>
                    )}
                    {editingStatsId === slide.id + 300 && (
                      <button
                        onClick={() => setEditingStatsId(null)}
                        style={{
                          padding: "6px 14px",
                          fontSize: "12px",
                          fontWeight: "600",
                          backgroundColor: "#2DAD70",
                          color: "#fff",
                          border: "none",
                          borderRadius: "4px",
                          cursor: "pointer",
                        }}
                      >
                        ✓ Done
                      </button>
                    )}
                    <button
                      onClick={() => openSlideImport(slide.id)}
                      style={{
                        padding: "6px 14px",
                        fontSize: "12px",
                        fontWeight: "600",
                        backgroundColor: "#2DAD70",
                        color: "#fff",
                        border: "none",
                        borderRadius: "4px",
                        cursor: "pointer",
                      }}
                    >
                      🔄 Update Stats
                    </button>
                  </div>
                )}
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(2, 1fr)", gap: "16px" }}>
                {Object.entries(slide.data.lifetime).map(([key, value]) => (
                  <div
                    key={key}
                    style={{
                      padding: "20px",
                      backgroundColor: "#EBF3FD",
                      borderRadius: "8px",
                      textAlign: "center",
                      border: "1px solid #4682E1",
                    }}
                  >
                    <div style={{ fontSize: "36px", fontWeight: "700", color: "#1863DC", marginBottom: "8px" }}>
                      {editingStatsId === slide.id + 300 ? (
                        <input
                          type="number"
                          defaultValue={Number(value)}
                          onBlur={(e) => updateSlideData(slide.id, ["lifetime", key], e.target.value)}
                          style={{
                            width: "100%",
                            fontSize: "32px",
                            padding: "4px",
                            border: "2px solid #1863DC",
                            borderRadius: "4px",
                            textAlign: "center",
                          }}
                        />
                      ) : (
                        <span>{String(value)}</span>
                      )}
                    </div>
                    <div
                      style={{
                        fontSize: "11px",
                        color: "#5A6872",
                        textTransform: "uppercase",
                        fontWeight: "600",
                        letterSpacing: "0.5px",
                      }}
                    >
                      {key === "paid" ? "Paid Signups" : key === "rating" ? "Rating" : key}
                    </div>
                  </div>
                ))}
              </div>
            </div>
          </div>
        );

      case "subscriptions":
        return (
          <div>
            {isEditMode && (
              <div style={{ marginBottom: "16px" }}>
                <button
                  onClick={() => {
                    const newRow = {
                      channel: "New Channel",
                      totalTarget: 0,
                      targetAsOnDate: 0,
                      actual: 0,
                      percentage: 0,
                    };
                    updateSlideData(slide.id, ["rows"], [...slide.data.rows, newRow]);
                  }}
                  style={{
                    padding: "8px 16px",
                    fontSize: "13px",
                    fontWeight: "600",
                    backgroundColor: "#2DAD70",
                    color: "#fff",
                    border: "none",
                    borderRadius: "4px",
                    cursor: "pointer",
                  }}
                >
                  + Add Row
                </button>
              </div>
            )}
            <table
              style={{ width: "100%", borderCollapse: "collapse", fontSize: "14px", fontFamily: "Inter, sans-serif" }}
            >
              <thead>
                <tr style={{ backgroundColor: "#F8FAFB", borderBottom: "2px solid #1863DC" }}>
                  <th style={{ padding: "12px 16px", textAlign: "left", fontWeight: "600", color: "#212121" }}>
                    Channel
                  </th>
                  <th style={{ padding: "12px 16px", textAlign: "left", fontWeight: "600", color: "#212121" }}>
                    Total Target
                  </th>
                  <th style={{ padding: "12px 16px", textAlign: "left", fontWeight: "600", color: "#212121" }}>
                    Target as on Date
                  </th>
                  <th style={{ padding: "12px 16px", textAlign: "left", fontWeight: "600", color: "#212121" }}>
                    Actual
                  </th>
                  <th style={{ padding: "12px 16px", textAlign: "left", fontWeight: "600", color: "#212121" }}>
                    % (Target as on Date)
                  </th>
                  {isEditMode && <th style={{ padding: "12px 16px" }}>Actions</th>}
                </tr>
              </thead>
              <tbody>
                {slide.data.rows.map((row, idx) => (
                  <tr
                    key={idx}
                    style={{ backgroundColor: idx % 2 ? "#F8FAFB" : "#fff", borderBottom: "1px solid #EAEEF2" }}
                  >
                    <td style={{ padding: "12px 16px", fontWeight: "600", color: "#212121" }}>
                      {isEditMode ? (
                        <input
                          key={`channel-${idx}`}
                          defaultValue={row.channel}
                          onBlur={(e) => updateSlideData(slide.id, ["rows", idx, "channel"], e.target.value)}
                          style={{
                            width: "100%",
                            padding: "6px 8px",
                            border: "1px solid #DBDFE4",
                            borderRadius: "4px",
                          }}
                        />
                      ) : (
                        row.channel
                      )}
                    </td>
                    <td style={{ padding: "12px 16px" }}>
                      {isEditMode ? (
                        <input
                          key={`target-${idx}`}
                          type="number"
                          defaultValue={row.totalTarget ?? ""}
                          onBlur={(e) => updateSlideData(slide.id, ["rows", idx, "totalTarget"], e.target.value)}
                          style={{
                            width: "100px",
                            padding: "6px 8px",
                            border: "1px solid #DBDFE4",
                            borderRadius: "4px",
                          }}
                        />
                      ) : (
                        row.totalTarget || "-"
                      )}
                    </td>
                    <td style={{ padding: "12px 16px" }}>
                      {isEditMode ? (
                        <input
                          key={`targetdate-${idx}`}
                          type="number"
                          defaultValue={row.targetAsOnDate ?? ""}
                          onBlur={(e) => updateSlideData(slide.id, ["rows", idx, "targetAsOnDate"], e.target.value)}
                          style={{
                            width: "100px",
                            padding: "6px 8px",
                            border: "1px solid #DBDFE4",
                            borderRadius: "4px",
                          }}
                        />
                      ) : (
                        row.targetAsOnDate || "-"
                      )}
                    </td>
                    <td style={{ padding: "12px 16px" }}>
                      {isEditMode ? (
                        <input
                          key={`actual-${idx}`}
                          type="number"
                          defaultValue={row.actual ?? ""}
                          onBlur={(e) => updateSlideData(slide.id, ["rows", idx, "actual"], e.target.value)}
                          style={{
                            width: "100px",
                            padding: "6px 8px",
                            border: "1px solid #DBDFE4",
                            borderRadius: "4px",
                          }}
                        />
                      ) : (
                        row.actual || "-"
                      )}
                    </td>
                    <td style={{ padding: "12px 16px" }}>
                      {isEditMode ? (
                        <input
                          key={`percentage-${idx}`}
                          type="number"
                          defaultValue={row.percentage ?? ""}
                          onBlur={(e) => updateSlideData(slide.id, ["rows", idx, "percentage"], e.target.value)}
                          style={{
                            width: "80px",
                            padding: "6px 8px",
                            border: "1px solid #DBDFE4",
                            borderRadius: "4px",
                          }}
                        />
                      ) : row.percentage !== null && row.percentage !== 0 ? (
                        `${row.percentage}%`
                      ) : (
                        "-"
                      )}
                    </td>
                    {isEditMode && (
                      <td style={{ padding: "12px 16px" }}>
                        <button
                          onClick={() => {
                            const newRows = slide.data.rows.filter((_, i) => i !== idx);
                            updateSlideData(slide.id, ["rows"], newRows);
                          }}
                          style={{
                            padding: "6px 12px",
                            fontSize: "12px",
                            fontWeight: "500",
                            backgroundColor: "#DC2143",
                            color: "#fff",
                            border: "none",
                            borderRadius: "4px",
                            cursor: "pointer",
                          }}
                        >
                          Delete
                        </button>
                      </td>
                    )}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        );

      default:
        return null;
    }
  };

  return (
    <div
      style={{
        padding: "24px",
        fontFamily: "Inter, -apple-system, BlinkMacSystemFont, sans-serif",
        backgroundColor: "#F3F5F7",
        minHeight: "100vh",
        overflowAnchor: "none",
      }}
    >
      <style>{`
        * {
          overflow-anchor: none;
        }
        input[type=number]::-webkit-inner-spin-button,
        input[type=number]::-webkit-outer-spin-button {
          -webkit-appearance: none;
          margin: 0;
        }
        input[type=number] {
          -moz-appearance: textfield;
          appearance: textfield;
        }
        @media print {
          body { 
            background: white !important;
            margin: 0;
            padding: 0;
          }
          .no-print { display: none !important; }
          .slide-section { 
            page-break-after: always;
            page-break-inside: avoid;
            box-shadow: none !important;
            margin: 0 !important;
            padding: 40px !important;
            border: none !important;
            background: white !important;
          }
          .slide-section:last-child {
            page-break-after: auto;
          }
          table {
            page-break-inside: avoid;
            width: 100% !important;
          }
          h1, h2, h3 {
            page-break-after: avoid;
            color: #212121 !important;
          }
          * {
            -webkit-print-color-adjust: exact !important;
            print-color-adjust: exact !important;
            color-adjust: exact !important;
          }
        }
      `}</style>

      {!isPresentMode && (
        <div
          className="no-print"
          style={{
            backgroundColor: "#fff",
            padding: "24px",
            marginBottom: "24px",
            borderRadius: "8px",
            boxShadow: "0 1px 3px rgba(0,0,0,0.08)",
            position: "sticky",
            top: 0,
            zIndex: 100,
          }}
        >
          <div
            style={{
              display: "flex",
              justifyContent: "space-between",
              alignItems: "center",
              marginBottom: "20px",
              flexWrap: "wrap",
              gap: "12px",
            }}
          >
            <div style={{ display: "flex", alignItems: "center", gap: "20px", flexWrap: "wrap" }}>
              <h1 style={{ fontSize: "28px", fontWeight: "700", color: "#212121", margin: 0 }}>Sprint Dashboard</h1>
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: "8px",
                  padding: "8px 16px",
                  backgroundColor: "#EBF3FD",
                  borderRadius: "6px",
                  border: "2px solid #1863DC",
                }}
              >
                <span style={{ fontSize: "14px", fontWeight: "600", color: "#5A6872" }}>Current Sprint:</span>
                {isEditMode ? (
                  <input
                    type="number"
                    value={currentSprint}
                    onChange={(e) => setCurrentSprint(Number(e.target.value))}
                    style={{
                      width: "80px",
                      padding: "4px 8px",
                      fontSize: "16px",
                      fontWeight: "700",
                      color: "#1863DC",
                      border: "1px solid #1863DC",
                      borderRadius: "4px",
                      textAlign: "center",
                    }}
                  />
                ) : (
                  <span style={{ fontSize: "20px", fontWeight: "700", color: "#1863DC" }}>{currentSprint}</span>
                )}
              </div>
            </div>
            <div style={{ display: "flex", gap: "12px", flexWrap: "wrap" }}>
              <button
                onClick={() => setIsEditMode(!isEditMode)}
                style={{
                  padding: "12px 20px",
                  fontSize: "14px",
                  fontWeight: "600",
                  border: "none",
                  borderRadius: "6px",
                  backgroundColor: isEditMode ? "#2DAD70" : "#1863DC",
                  color: "#fff",
                  cursor: "pointer",
                  boxShadow: "0 2px 4px rgba(0,0,0,0.1)",
                }}
              >
                {isEditMode ? "✓ Save" : "✎ Edit"}
              </button>
              <button
                onClick={openGlobalImport}
                style={{
                  padding: "12px 20px",
                  fontSize: "14px",
                  fontWeight: "600",
                  border: "none",
                  borderRadius: "6px",
                  backgroundColor: "#2DAD70",
                  color: "#fff",
                  cursor: "pointer",
                  boxShadow: "0 2px 4px rgba(0,0,0,0.1)",
                }}
              >
                📊 Import Data
              </button>
              <button
                onClick={exportToJSON}
                style={{
                  padding: "12px 20px",
                  fontSize: "14px",
                  fontWeight: "600",
                  border: "none",
                  borderRadius: "6px",
                  backgroundColor: "#FF8800",
                  color: "#fff",
                  cursor: "pointer",
                  boxShadow: "0 2px 4px rgba(0,0,0,0.1)",
                }}
              >
                💾 Export Data
              </button>
              <button
                onClick={exportToPDF}
                style={{
                  padding: "12px 20px",
                  fontSize: "14px",
                  fontWeight: "600",
                  border: "none",
                  borderRadius: "6px",
                  backgroundColor: "#7F56D9",
                  color: "#fff",
                  cursor: "pointer",
                  boxShadow: "0 2px 4px rgba(0,0,0,0.1)",
                }}
              >
                📄 Export PDF
              </button>
              <button
                onClick={enterPresentMode}
                style={{
                  padding: "12px 20px",
                  fontSize: "14px",
                  fontWeight: "600",
                  border: "none",
                  borderRadius: "6px",
                  backgroundColor: "#363F52",
                  color: "#fff",
                  cursor: "pointer",
                  boxShadow: "0 2px 4px rgba(0,0,0,0.1)",
                }}
              >
                🎬 Present
              </button>
            </div>
          </div>

          <div
            style={{
              backgroundColor: "#FEF9E6",
              border: "1px solid #FFB800",
              borderRadius: "6px",
              padding: "16px",
              fontSize: "13px",
              color: "#5A4A00",
              marginBottom: "12px",
            }}
          >
            <strong>💡 Edit Mode:</strong> Edit slide titles, add/remove rows & columns, edit all values, rename
            headers. Latest sprint row is bold with a blue border.
            <br />
            <strong>📊 Import Data:</strong> Upload CSV or paste data to auto-fill. Maximum 5 sprints will be kept
            (oldest removed automatically).
          </div>
        </div>
      )}

      {showImportModal && (
        <div
          style={{
            position: "fixed",
            top: 0,
            left: 0,
            right: 0,
            bottom: 0,
            backgroundColor: "rgba(0,0,0,0.5)",
            zIndex: 2000,
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
          }}
        >
          <div
            style={{
              backgroundColor: "#fff",
              borderRadius: "12px",
              padding: "32px",
              maxWidth: "600px",
              width: "90%",
              boxShadow: "0 8px 32px rgba(0,0,0,0.3)",
              position: "relative",
            }}
          >
            <button
              onClick={() => {
                setShowImportModal(false);
                setCurrentImportSlideId(null);
                setPastedData("");
                setPastedImage(null);
                setIsGlobalImport(false);
              }}
              style={{
                position: "absolute",
                top: "16px",
                right: "16px",
                width: "32px",
                height: "32px",
                border: "none",
                borderRadius: "6px",
                backgroundColor: "#F1F3F5",
                color: "#5A6872",
                cursor: "pointer",
                fontSize: "18px",
                fontWeight: "700",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                transition: "all 0.2s",
              }}
              onMouseEnter={(e) => {
                e.currentTarget.style.backgroundColor = "#E1E4E8";
              }}
              onMouseLeave={(e) => {
                e.currentTarget.style.backgroundColor = "#F1F3F5";
              }}
            >
              ×
            </button>
            <h2
              style={{
                fontSize: "24px",
                fontWeight: "700",
                color: "#212121",
                marginBottom: "20px",
                paddingRight: "40px",
              }}
            >
              {isGlobalImport ? "Import Full Presentation" : "Update Slide Data"}
            </h2>
            <p style={{ fontSize: "14px", color: "#5A6872", marginBottom: "20px" }}>
              {isGlobalImport
                ? "Upload a JSON export to restore complete data, or paste CSV format below to bulk update all slides at once."
                : "Copy data from your Google Sheet and paste it below, OR paste/upload a screenshot to extract data automatically."}
            </p>

            {isGlobalImport && (
              <div style={{ marginBottom: "20px" }}>
                <label
                  style={{
                    display: "inline-block",
                    padding: "12px 24px",
                    fontSize: "14px",
                    fontWeight: "600",
                    backgroundColor: "#2DAD70",
                    color: "#fff",
                    borderRadius: "6px",
                    cursor: "pointer",
                    marginBottom: "16px",
                  }}
                >
                  📁 Select JSON File
                  <input
                    type="file"
                    accept=".json,application/json"
                    onChange={handleFileImport}
                    style={{ display: "none" }}
                  />
                </label>

                <div
                  style={{
                    backgroundColor: "#F8F9FA",
                    border: "1px solid #DFE3E8",
                    borderRadius: "8px",
                    padding: "16px",
                    fontSize: "13px",
                    color: "#212121",
                    fontFamily: "monospace",
                    lineHeight: "1.8",
                  }}
                >
                  <div style={{ fontWeight: "700", fontSize: "14px", marginBottom: "12px", color: "#1863DC" }}>
                    📋 Bulk Paste Format (CSV):
                  </div>
                  <div style={{ marginBottom: "12px", color: "#5A6872" }}>
                    Paste data in this format to update all slides at once:
                  </div>
                  <div
                    style={{
                      backgroundColor: "#fff",
                      padding: "12px",
                      borderRadius: "4px",
                      border: "1px solid #E1E4E8",
                      overflowX: "auto",
                    }}
                  >
                    <strong>SlideID,TableName,Sprint,[Column Names...]</strong>
                    <br />
                    1,positionChanges,263,15,45
                    <br />
                    1,positionChanges,264,18,42
                    <br />
                    2,main,263,5,3,8,2
                    <br />
                    7,tickets,263,350,2h 30m,5h 45m,95%
                    <br />
                    16,main,263,350,25,325,120,15
                  </div>
                  <div style={{ marginTop: "12px", fontSize: "12px", color: "#5A6872" }}>
                    <strong>Key Points:</strong>
                    <br />• <strong>SlideID:</strong> The slide number (1-16)
                    <br />• <strong>TableName:</strong> "main", "positionChanges", "tickets", "liveChat", etc.
                    <br />• <strong>Sprint:</strong> Sprint number (e.g., 263, 264)
                    <br />• <strong>Columns:</strong> Match your slide's column names exactly
                    <br />• Each row updates one sprint in one table
                  </div>
                </div>
              </div>
            )}

            {!isGlobalImport && (
              <>
                <div
                  style={{
                    marginBottom: "20px",
                    padding: "16px",
                    backgroundColor: "#FEF3E6",
                    borderRadius: "6px",
                    border: "1px solid #FF8800",
                  }}
                >
                  <div style={{ fontSize: "13px", color: "#994D00", marginBottom: "8px" }}>
                    <strong>📸 Screenshot Upload (NEW!):</strong>
                  </div>
                  <ul style={{ margin: 0, paddingLeft: "20px", fontSize: "13px", color: "#994D00", lineHeight: "1.6" }}>
                    <li>Take a screenshot of your Google Sheet or table</li>
                    <li>Paste directly (Ctrl+V) in the textarea below OR use the upload button</li>
                    <li>AI will automatically extract and format the data for you</li>
                  </ul>
                </div>

                <div
                  style={{
                    marginBottom: "20px",
                    padding: "16px",
                    backgroundColor: "#E6F7ED",
                    borderRadius: "6px",
                    border: "1px solid #2DAD70",
                  }}
                >
                  <div style={{ fontSize: "13px", color: "#1D7A47", marginBottom: "8px" }}>
                    <strong>📊 For Tables (Text Paste):</strong>
                  </div>
                  <ul style={{ margin: 0, paddingLeft: "20px", fontSize: "13px", color: "#1D7A47", lineHeight: "1.6" }}>
                    <li>
                      <strong>Step 1:</strong> In Google Sheets, select your entire table (header + data rows)
                    </li>
                    <li>
                      <strong>Step 2:</strong> Copy (Ctrl+C or Cmd+C)
                    </li>
                    <li>
                      <strong>Step 3:</strong> Paste here - data will be tab-separated automatically
                    </li>
                    <li>Keeps last 5 sprints automatically | Numbers cleaned from $, %, commas</li>
                  </ul>
                </div>

                <div
                  style={{
                    marginBottom: "20px",
                    padding: "16px",
                    backgroundColor: "#EBF3FD",
                    borderRadius: "6px",
                    border: "1px solid #1863DC",
                  }}
                >
                  <div style={{ fontSize: "13px", color: "#134FB0", marginBottom: "8px" }}>
                    <strong>📈 For Stats (3 supported formats):</strong>
                  </div>
                  <ul style={{ margin: 0, paddingLeft: "20px", fontSize: "13px", color: "#134FB0", lineHeight: "1.6" }}>
                    <li>
                      <strong>Format 1:</strong> Key-value with tab: <code>total&nbsp;&nbsp;&nbsp;&nbsp;150</code>
                    </li>
                    <li>
                      <strong>Format 2:</strong> Key-value with colon: <code>Total: 150</code>
                    </li>
                    <li>
                      <strong>Format 3:</strong> Just values (one per line in order)
                    </li>
                  </ul>
                </div>

                <div style={{ marginBottom: "20px" }}>
                  <label
                    style={{
                      display: "block",
                      marginBottom: "8px",
                      fontSize: "14px",
                      fontWeight: "600",
                      color: "#212121",
                    }}
                  >
                    Paste Data or Screenshot Here
                  </label>
                  {isProcessingImage && (
                    <div
                      style={{
                        padding: "12px",
                        backgroundColor: "#EBF3FD",
                        borderRadius: "6px",
                        marginBottom: "12px",
                        fontSize: "13px",
                        color: "#1863DC",
                        textAlign: "center",
                      }}
                    >
                      🔄 Processing screenshot with AI... This may take a few seconds.
                    </div>
                  )}
                  {pastedImage && !isProcessingImage && (
                    <div style={{ marginBottom: "12px" }}>
                      <img
                        src={pastedImage}
                        alt="Pasted screenshot"
                        style={{
                          maxWidth: "100%",
                          maxHeight: "200px",
                          borderRadius: "6px",
                          border: "2px solid #2DAD70",
                        }}
                      />
                    </div>
                  )}
                  <textarea
                    value={pastedData}
                    onChange={(e) => setPastedData(e.target.value)}
                    onPaste={handlePasteEvent}
                    placeholder="Paste your data OR screenshot here (Ctrl+V)...&#10;&#10;Text: Sprint  Total  Social&#10;      263     45     12&#10;&#10;Screenshot: Just paste from clipboard!"
                    style={{
                      width: "100%",
                      minHeight: "150px",
                      padding: "12px",
                      border: "2px solid #DBDFE4",
                      borderRadius: "6px",
                      fontSize: "13px",
                      fontFamily: "monospace",
                      resize: "vertical",
                    }}
                    disabled={isProcessingImage}
                  />
                </div>

                <div style={{ display: "flex", gap: "12px", justifyContent: "flex-end" }}>
                  <button
                    onClick={() => {
                      setShowImportModal(false);
                      setCurrentImportSlideId(null);
                      setPastedData("");
                      setPastedImage(null);
                      setIsGlobalImport(false);
                    }}
                    style={{
                      padding: "12px 24px",
                      fontSize: "14px",
                      fontWeight: "600",
                      border: "2px solid #DBDFE4",
                      borderRadius: "6px",
                      backgroundColor: "#fff",
                      color: "#5A6872",
                      cursor: "pointer",
                    }}
                  >
                    Cancel
                  </button>
                  {!isGlobalImport && (
                    <button
                      onClick={handlePastedData}
                      disabled={isProcessingImage}
                      style={{
                        padding: "12px 24px",
                        fontSize: "14px",
                        fontWeight: "600",
                        border: "none",
                        borderRadius: "6px",
                        backgroundColor: isProcessingImage ? "#DBDFE4" : "#2DAD70",
                        color: "#fff",
                        cursor: isProcessingImage ? "not-allowed" : "pointer",
                        opacity: isProcessingImage ? 0.6 : 1,
                      }}
                    >
                      📥 Import Data
                    </button>
                  )}
                </div>
              </>
            )}
          </div>
        </div>
      )}

      {isPresentMode && (
        <div
          style={{
            position: "fixed",
            top: 0,
            left: 0,
            right: 0,
            bottom: 0,
            backgroundColor: "#1C2630",
            zIndex: 1000,
            display: "flex",
            flexDirection: "column",
            overflow: "auto",
          }}
        >
          <div
            style={{
              position: "fixed",
              top: "24px",
              right: "24px",
              zIndex: 1001,
              display: "flex",
              gap: "12px",
              alignItems: "center",
            }}
          >
            <div
              style={{
                color: "#fff",
                fontSize: "16px",
                fontWeight: "500",
                backgroundColor: "rgba(0,0,0,0.3)",
                padding: "8px 16px",
                borderRadius: "6px",
              }}
            >
              {currentSlide} / {sprintData.slides.length}
            </div>
            <button
              onClick={exitPresentMode}
              style={{
                padding: "12px 24px",
                backgroundColor: "#DC2143",
                color: "#fff",
                border: "none",
                borderRadius: "6px",
                cursor: "pointer",
                fontWeight: "600",
                fontSize: "14px",
              }}
            >
              ✕ Exit (Esc)
            </button>
          </div>

          <div
            style={{
              flex: 1,
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              padding: "80px 40px 40px 40px",
              position: "relative",
            }}
          >
            <button
              onClick={goToPreviousSlide}
              disabled={currentSlide === 1}
              style={{
                position: "absolute",
                left: "20px",
                top: "50%",
                transform: "translateY(-50%)",
                padding: "16px 20px",
                backgroundColor: currentSlide === 1 ? "#5A6872" : "#1863DC",
                color: "#fff",
                border: "none",
                borderRadius: "8px",
                cursor: currentSlide === 1 ? "not-allowed" : "pointer",
                fontWeight: "600",
                fontSize: "18px",
                opacity: currentSlide === 1 ? 0.5 : 1,
                zIndex: 10,
              }}
            >
              ←
            </button>

            <div
              style={{
                maxWidth: "1400px",
                width: "100%",
                backgroundColor: "#fff",
                borderRadius: "12px",
                boxShadow: "0 8px 32px rgba(0,0,0,0.3)",
                transformOrigin: "center center",
                transform: "scale(0.85)",
              }}
            >
              {sprintData.slides[currentSlide - 1] && (
                <div style={{ padding: "48px" }}>
                  {currentSprint > 0 && (
                    <div
                      style={{
                        display: "inline-flex",
                        alignItems: "center",
                        gap: "8px",
                        padding: "10px 20px",
                        backgroundColor: "#EBF3FD",
                        borderRadius: "8px",
                        border: "2px solid #1863DC",
                        marginBottom: "24px",
                      }}
                    >
                      <span style={{ fontSize: "16px", fontWeight: "600", color: "#5A6872" }}>Sprint:</span>
                      <span style={{ fontSize: "24px", fontWeight: "700", color: "#1863DC" }}>{currentSprint}</span>
                    </div>
                  )}
                  <div
                    style={{
                      display: "flex",
                      justifyContent: "space-between",
                      alignItems: "center",
                      marginBottom: "32px",
                      borderBottom: "3px solid #1863DC",
                      paddingBottom: "16px",
                    }}
                  >
                    <h2 style={{ fontSize: "28px", fontWeight: "700", color: "#212121", margin: 0 }}>
                      {sprintData.slides[currentSlide - 1].title}
                    </h2>
                    {sprintData.slides[currentSlide - 1].moreDetailsUrl && (
                      <a
                        href={sprintData.slides[currentSlide - 1].moreDetailsUrl}
                        target="_blank"
                        rel="noopener noreferrer"
                        style={{
                          padding: "10px 20px",
                          fontSize: "14px",
                          fontWeight: "600",
                          backgroundColor: "#EBF3FD",
                          color: "#1863DC",
                          textDecoration: "none",
                          borderRadius: "24px",
                          whiteSpace: "nowrap",
                          border: "1px solid #1863DC",
                        }}
                      >
                        📊 More Details
                      </a>
                    )}
                  </div>
                  {renderSlideContent(sprintData.slides[currentSlide - 1])}
                </div>
              )}
            </div>

            <button
              onClick={goToNextSlide}
              disabled={currentSlide === sprintData.slides.length}
              style={{
                position: "absolute",
                right: "20px",
                top: "50%",
                transform: "translateY(-50%)",
                padding: "16px 20px",
                backgroundColor: currentSlide === sprintData.slides.length ? "#5A6872" : "#1863DC",
                color: "#fff",
                border: "none",
                borderRadius: "8px",
                cursor: currentSlide === sprintData.slides.length ? "not-allowed" : "pointer",
                fontWeight: "600",
                fontSize: "18px",
                opacity: currentSlide === sprintData.slides.length ? 0.5 : 1,
                zIndex: 10,
              }}
            >
              →
            </button>
          </div>

          <div
            style={{
              padding: "20px",
              textAlign: "center",
              color: "#fff",
              fontSize: "13px",
              backgroundColor: "rgba(0,0,0,0.2)",
            }}
          >
            Use ← → arrow keys, Space, or on-screen buttons to navigate | Press Esc to exit
          </div>
        </div>
      )}

      <div style={{ maxWidth: "1400px", margin: "0 auto", display: isPresentMode ? "none" : "block" }}>
        {sprintData.slides.map((slide) => (
          <Slide key={slide.id} slide={slide} />
        ))}
      </div>
    </div>
  );
}
