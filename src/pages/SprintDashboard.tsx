import { useState, useEffect } from 'react';

export default function SprintDashboard() {
  const [isEditMode, setIsEditMode] = useState(false);
  const [isPresentMode, setIsPresentMode] = useState(false);
  const [currentSlide, setCurrentSlide] = useState(1);
  const [currentSprint, setCurrentSprint] = useState(0);
  const [showImportModal, setShowImportModal] = useState(false);
  const [currentImportSlideId, setCurrentImportSlideId] = useState<number | null>(null);

  const MAX_SPRINTS = 5;

  const maintainSprintLimit = (rows) => {
    if (!rows || rows.length <= MAX_SPRINTS) return rows;
    
    // Sort by sprint number (descending) and keep only last 5
    const sorted = [...rows].sort((a, b) => {
      const sprintA = typeof a.sprint === 'number' ? a.sprint : parseInt(a.sprint) || 0;
      const sprintB = typeof b.sprint === 'number' ? b.sprint : parseInt(b.sprint) || 0;
      return sprintB - sprintA;
    });
    
    return sorted.slice(0, MAX_SPRINTS);
  };

  const handleFileUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const text = e.target?.result as string;
        const lines = text.split('\n').filter(row => row.trim());
        
        if (lines.length < 2) {
          alert('CSV file must contain headers and at least one data row.');
          return;
        }

        // Parse CSV
        const headers = lines[0].split(',').map(h => h.trim().toLowerCase());
        const dataRows = lines.slice(1).map(line => {
          const values = line.split(',').map(v => v.trim());
          const row: any = {};
          headers.forEach((header, idx) => {
            const value = values[idx] || '';
            // Try to convert to number if possible
            row[header] = isNaN(value as any) || value === '' ? value : Number(value);
          });
          return row;
        });

        // Apply 5-sprint limit - keep only last 5 rows
        const limitedRows = maintainSprintLimit(dataRows);

        // Update the specific slide with imported data
        if (currentImportSlideId) {
          importDataToSlide(currentImportSlideId, headers, limitedRows);
        }

        alert(`Data imported successfully! ${limitedRows.length} sprint(s) loaded (max 5 sprints maintained).`);
        setShowImportModal(false);
        setCurrentImportSlideId(null);
      } catch (error) {
        console.error('Import error:', error);
        alert('Error reading file. Please check the CSV format and try again.');
      }
    };
    reader.readAsText(file);
  };

  const importDataToSlide = (slideId, headers, rows) => {
    setSprintData(prev => ({
      ...prev,
      slides: prev.slides.map(slide => {
        if (slide.id !== slideId) return slide;

        const newSlide = JSON.parse(JSON.stringify(slide));

        // Handle different slide types
        if (slide.type === 'table' || slide.type === 'pluginRanking') {
          // Direct table import
          if (newSlide.data.rows) {
            newSlide.data.rows = rows;
          }
        } else if (slide.type === 'supportData') {
          // Check if importing tickets or live chat based on headers
          if (headers.includes('totaltickets') || headers.includes('total tickets solved')) {
            newSlide.data.tickets.rows = rows;
          } else if (headers.includes('conversations') || headers.includes('conversations assigned')) {
            newSlide.data.liveChat.rows = rows;
          }
        } else if (slide.type === 'agencyLeads') {
          // Check if importing leads conversion or Q3 performance
          if (headers.includes('metrics')) {
            newSlide.data.leadsConversion.rows = rows;
          } else if (headers.includes('quarter')) {
            newSlide.data.q3Performance.rows = rows;
          }
        } else if (slide.type === 'quarterStats') {
          // Import sprint data and quarter stats
          const sprintRows = rows.filter(r => r.sprint);
          if (sprintRows.length > 0) {
            newSlide.data.rows = sprintRows;
          }
          // Update quarter stats if present
          if (rows[0] && !rows[0].sprint) {
            const statsRow = rows[0];
            Object.keys(newSlide.data.quarterStats).forEach(key => {
              if (statsRow[key] !== undefined) {
                newSlide.data.quarterStats[key] = statsRow[key];
              }
            });
          }
        } else if (slide.type === 'withTarget') {
          newSlide.data.rows = rows;
          // Update totals if present
          const totalsRow = rows.find(r => r.sprint === 'total' || r.sprint === 'qtd');
          if (totalsRow && newSlide.data.total) {
            Object.keys(newSlide.data.total).forEach(key => {
              if (totalsRow[key] !== undefined) {
                newSlide.data.total[key] = totalsRow[key];
              }
            });
          }
        } else if (slide.type === 'referral') {
          const sprintRows = rows.filter(r => typeof r.sprint === 'number' || !isNaN(parseInt(r.sprint)));
          newSlide.data.rows = sprintRows;
          
          // Update lifetime stats if present
          const lifetimeRow = rows.find(r => r.sprint === 'lifetime' || r.type === 'lifetime');
          if (lifetimeRow && newSlide.data.lifetime) {
            Object.keys(newSlide.data.lifetime).forEach(key => {
              if (lifetimeRow[key] !== undefined) {
                newSlide.data.lifetime[key] = lifetimeRow[key];
              }
            });
          }
        } else if (slide.type === 'wixApp') {
          const sprintRows = rows.filter(r => typeof r.sprint === 'number' || !isNaN(parseInt(r.sprint)));
          newSlide.data.rows = sprintRows;
          
          // Update lifetime stats
          const lifetimeRow = rows.find(r => r.sprint === 'lifetime' || r.type === 'lifetime');
          if (lifetimeRow && newSlide.data.lifetime) {
            Object.keys(newSlide.data.lifetime).forEach(key => {
              if (lifetimeRow[key] !== undefined) {
                newSlide.data.lifetime[key] = lifetimeRow[key];
              }
            });
          }
        } else if (slide.type === 'subscriptions') {
          newSlide.data.rows = rows;
        } else if (slide.type === 'rankings') {
          // Update position changes table
          const sprintRows = rows.filter(r => r.sprint);
          if (sprintRows.length > 0) {
            newSlide.data.positionChanges.rows = sprintRows;
          }
        }

        return newSlide;
      })
    }));
  };

  const openSlideImport = (slideId) => {
    const slide = sprintData.slides.find(s => s.id === slideId);
    if (!slide.moreDetailsUrl) {
      alert('Please add a Google Sheet URL in the "Sheet URL" field first, then use Update Slide to pull data from it.');
      return;
    }
    setCurrentImportSlideId(slideId);
    setShowImportModal(true);
  };
  
  const [sprintData, setSprintData] = useState({
    slides: [
      {
        id: 1,
        title: 'Top 25 Major Rankings and Movements',
        type: 'rankings',
        moreDetailsUrl: 'https://docs.google.com/spreadsheets',
        data: {
          total: 0,
          byRegion: [
            { region: 'US', count: 0, keywords: '' },
            { region: 'UK', count: 0, keywords: '' },
            { region: 'DE', count: 0, keywords: '' }
          ],
          positionChanges: {
            columns: [
              { key: 'sprint', header: 'Sprint', locked: true },
              { key: 'pos1_2', header: 'Position 1-2' },
              { key: 'pos3_10', header: 'Position 3-10' }
            ],
            rows: [
              { sprint: 263, pos1_2: 0, pos3_10: 0 },
              { sprint: 264, pos1_2: 0, pos3_10: 0 },
              { sprint: 265, pos1_2: 0, pos3_10: 0 }
            ]
          },
          improved: [
            { region: 'US', count: 0, keywords: '' }
          ],
          declined: [
            { region: 'US', count: 0, keywords: '' }
          ]
        }
      },
      {
        id: 2,
        title: 'Content Publishing Stats',
        type: 'table',
        moreDetailsUrl: 'https://docs.google.com/spreadsheets',
        data: {
          columns: [
            { key: 'sprint', header: 'Sprint', locked: true },
            { key: 'blog', header: 'Blog Posts' },
            { key: 'infographics', header: 'Infographics' },
            { key: 'kb', header: 'KB Articles' },
            { key: 'videos', header: 'Videos' }
          ],
          rows: [
            { sprint: 263, blog: 0, infographics: 0, kb: 0, videos: 0 },
            { sprint: 264, blog: 0, infographics: 0, kb: 0, videos: 0 },
            { sprint: 265, blog: 0, infographics: 0, kb: 0, videos: 0 }
          ]
        }
      },
      {
        id: 3,
        title: 'Brand Mentions',
        type: 'table',
        moreDetailsUrl: 'https://docs.google.com/spreadsheets',
        data: {
          columns: [
            { key: 'sprint', header: 'Sprint', locked: true },
            { key: 'total', header: 'Total' },
            { key: 'social', header: 'Social' },
            { key: 'blogs', header: 'Blogs' },
            { key: 'youtube', header: 'YouTube' },
            { key: 'negative', header: 'Negative' }
          ],
          rows: [
            { sprint: 263, total: 0, social: 0, blogs: 0, youtube: 0, negative: 0 },
            { sprint: 264, total: 0, social: 0, blogs: 0, youtube: 0, negative: 0 },
            { sprint: 265, total: 0, social: 0, blogs: 0, youtube: 0, negative: 0 }
          ]
        }
      },
      {
        id: 4,
        title: 'WP Popular Plugin Ranking',
        type: 'pluginRanking',
        moreDetailsUrl: 'https://docs.google.com/spreadsheets',
        data: {
          quarterTarget: 0,
          columns: [
            { key: 'sprint', header: 'Sprint', locked: true },
            { key: 'position', header: 'Plugin Position*' }
          ],
          rows: [
            { sprint: 263, position: 0 },
            { sprint: 264, position: 0 },
            { sprint: 265, position: 0 }
          ]
        }
      },
      {
        id: 5,
        title: 'Plugin Paid Connections',
        type: 'table',
        moreDetailsUrl: 'https://docs.google.com/spreadsheets',
        data: {
          columns: [
            { key: 'sprint', header: 'Sprint', locked: true },
            { key: 'total', header: 'Total Paid' },
            { key: 'direct', header: 'Direct Plans' }
          ],
          rows: [
            { sprint: 263, total: 0, direct: 0 },
            { sprint: 264, total: 0, direct: 0 },
            { sprint: 265, total: 0, direct: 0 }
          ]
        }
      },
      {
        id: 7,
        title: 'Support data',
        type: 'supportData',
        moreDetailsUrl: 'https://docs.google.com/spreadsheets',
        data: {
          tickets: {
            columns: [
              { key: 'sprint', header: 'Sprint', locked: true },
              { key: 'totalTickets', header: 'Total Tickets Solved' },
              { key: 'avgFirstResponse', header: 'Avg First Response Time' },
              { key: 'avgFullResolution', header: 'Avg Full Resolution time' },
              { key: 'csat', header: 'CSAT Score' },
              { key: 'presales', header: 'Pre-sales Tickets' },
              { key: 'converted', header: 'Converted Tickets (Unique customers)' },
              { key: 'paidSubs', header: 'Total Paid subscriptions (websites)' },
              { key: 'agencyTickets', header: 'Agency Tickets' },
              { key: 'badRating', header: 'Bad Rating' }
            ],
            rows: [
              { sprint: 263, totalTickets: 0, avgFirstResponse: '', avgFullResolution: '', csat: '', presales: 0, converted: 0, paidSubs: 0, agencyTickets: 0, badRating: '-' },
              { sprint: 264, totalTickets: 0, avgFirstResponse: '', avgFullResolution: '', csat: '', presales: 0, converted: 0, paidSubs: 0, agencyTickets: 0, badRating: '-' },
              { sprint: 265, totalTickets: 0, avgFirstResponse: '', avgFullResolution: '', csat: '', presales: 0, converted: 0, paidSubs: 0, agencyTickets: 0, badRating: '-' }
            ]
          },
          liveChat: {
            columns: [
              { key: 'sprint', header: 'Sprint', locked: true },
              { key: 'conversations', header: 'Conversations assigned' },
              { key: 'avgAssignment', header: 'Avg teammate assignment to first response' },
              { key: 'avgResolution', header: 'Avg full resolution time' },
              { key: 'csat', header: 'CSAT Score' },
              { key: 'badRating', header: 'Bad Rating' }
            ],
            rows: [
              { sprint: 263, conversations: 0, avgAssignment: '', avgResolution: '', csat: '', badRating: '-' },
              { sprint: 264, conversations: 0, avgAssignment: '', avgResolution: '', csat: '', badRating: '-' },
              { sprint: 265, conversations: 0, avgAssignment: '', avgResolution: '', csat: '', badRating: '-' }
            ]
          }
        }
      },
      {
        id: 8,
        title: 'Agency Leads & Conversion (Presales)',
        type: 'agencyLeads',
        moreDetailsUrl: 'https://docs.google.com/spreadsheets',
        data: {
          leadsConversion: {
            columns: [
              { key: 'metrics', header: 'Metrics', locked: true },
              { key: 'totalCount', header: 'Total Count' },
              { key: 'fromTickets', header: 'From Tickets' },
              { key: 'websiteLeads', header: 'Website Leads (LP)' },
              { key: 'fromAds', header: 'From Ads' },
              { key: 'liveChat', header: 'Live chat' },
              { key: 'webApp', header: 'Web app (Book a call)' }
            ],
            rows: [
              { metrics: 'Total Leads Received', totalCount: 0, fromTickets: 0, websiteLeads: 0, fromAds: 0, liveChat: 0, webApp: 0 },
              { metrics: 'Agency Demos', totalCount: 0, fromTickets: 0, websiteLeads: 0, fromAds: 0, liveChat: 0, webApp: 0 },
              { metrics: 'New Agency Signups', totalCount: 0, fromTickets: 0, websiteLeads: 0, fromAds: 0, liveChat: 0, webApp: 0 },
              { metrics: 'Paid Conversions', totalCount: 0, fromTickets: 0, websiteLeads: 0, fromAds: 0, liveChat: 0, webApp: 0 }
            ]
          },
          q3Performance: {
            columns: [
              { key: 'quarter', header: 'Quarter', locked: true },
              { key: 'target', header: 'Target' },
              { key: 'achieved', header: 'Achieved' },
              { key: 'percentage', header: 'Percentage' }
            ],
            rows: [
              { quarter: 'Sprint 263', target: 0, achieved: 0, percentage: '0%' },
              { quarter: 'Sprint 264', target: 0, achieved: 0, percentage: '0%' },
              { quarter: 'Sprint 265', target: 0, achieved: 0, percentage: '0%' }
            ]
          }
        }
      },
      {
        id: 9,
        title: 'Paid Acquisition - Google Ads',
        type: 'quarterStats',
        moreDetailsUrl: 'https://docs.google.com/spreadsheets',
        data: {
          quarterStats: {
            accountsCreated: 0,
            cardAdded: 0,
            bannerActive: 0,
            payingUsers: 0,
            roas: 0
          },
          columns: [
            { key: 'sprint', header: 'Sprint', locked: true },
            { key: 'totalAccounts', header: 'Total Accounts' },
            { key: 'paidTrials', header: 'Paid Trials' },
            { key: 'payingUsers', header: 'Paying Users' }
          ],
          rows: [
            { sprint: 263, totalAccounts: 0, paidTrials: 0, payingUsers: 0 },
            { sprint: 264, totalAccounts: 0, paidTrials: 0, payingUsers: 0 },
            { sprint: 265, totalAccounts: 0, paidTrials: 0, payingUsers: 0 }
          ]
        }
      },
      {
        id: 10,
        title: 'Key Google Ads Observations',
        type: 'textarea',
        moreDetailsUrl: 'https://docs.google.com/spreadsheets',
        data: {
          text: ''
        }
      },
      {
        id: 11,
        title: 'Paid Acquisition - Bing Ads',
        type: 'quarterStats',
        moreDetailsUrl: 'https://docs.google.com/spreadsheets',
        data: {
          quarterStats: {
            accountsCreated: 0,
            cardAdded: 0,
            bannerActive: 0,
            payingUsers: 0,
            roas: 0
          },
          columns: [
            { key: 'sprint', header: 'Sprint', locked: true },
            { key: 'totalAccounts', header: 'Total Accounts' },
            { key: 'paidTrials', header: 'Paid Trials' },
            { key: 'payingUsers', header: 'Paying Users' }
          ],
          rows: [
            { sprint: 263, totalAccounts: 0, paidTrials: 0, payingUsers: 0 },
            { sprint: 264, totalAccounts: 0, paidTrials: 0, payingUsers: 0 },
            { sprint: 265, totalAccounts: 0, paidTrials: 0, payingUsers: 0 }
          ]
        }
      },
      {
        id: 12,
        title: 'Paid Acquisition - Agency Data',
        type: 'table',
        moreDetailsUrl: 'https://docs.google.com/spreadsheets',
        data: {
          columns: [
            { key: 'sprint', header: 'Sprint', locked: true },
            { key: 'formFills', header: 'Form Fills' },
            { key: 'demos', header: 'Demo' },
            { key: 'signups', header: 'Signups' },
            { key: 'paying', header: 'Paying Agencies' }
          ],
          rows: [
            { sprint: 263, formFills: 0, demos: 0, signups: 0, paying: 0 },
            { sprint: 264, formFills: 0, demos: 0, signups: 0, paying: 0 },
            { sprint: 265, formFills: 0, demos: 0, signups: 0, paying: 0 }
          ]
        }
      },
      {
        id: 13,
        title: 'Agency - Signups & Paid Users',
        type: 'withTarget',
        moreDetailsUrl: 'https://docs.google.com/spreadsheets',
        data: {
          target: 0,
          columns: [
            { key: 'sprint', header: 'Sprint', locked: true },
            { key: 'signups', header: 'Signups' },
            { key: 'paid', header: 'Paid' },
            { key: 'percentage', header: 'Target Achieved %' },
            { key: 'shortfall', header: 'Shortfall' }
          ],
          rows: [
            { sprint: 263, signups: 0, paid: 0, percentage: 0, shortfall: 0 },
            { sprint: 264, signups: 0, paid: 0, percentage: 0, shortfall: 0 },
            { sprint: 265, signups: 0, paid: 0, percentage: 0, shortfall: 0 }
          ],
          total: { signups: 0, paid: 0 }
        }
      },
      {
        id: 14,
        title: 'Partnerships & Growth - Affiliate Partner Program',
        type: 'withTarget',
        moreDetailsUrl: 'https://docs.google.com/spreadsheets',
        data: {
          target: 0,
          columns: [
            { key: 'sprint', header: 'Sprint', locked: true },
            { key: 'newAff', header: 'New Affiliates' },
            { key: 'trials', header: 'Trial Signups' },
            { key: 'paid', header: 'Paid Signups' }
          ],
          rows: [
            { sprint: 263, newAff: 0, trials: 0, paid: 0 },
            { sprint: 264, newAff: 0, trials: 0, paid: 0 },
            { sprint: 265, newAff: 0, trials: 0, paid: 0 }
          ],
          total: { newAff: 0, trials: 0, paid: 0 }
        }
      },
      {
        id: 15,
        title: 'Partnerships & Growth - Referral Partner Program',
        type: 'referral',
        moreDetailsUrl: 'https://docs.google.com/spreadsheets',
        data: {
          lifetime: { advocates: 0, paid: 0 },
          target: 0,
          current: 0,
          columns: [
            { key: 'sprint', header: 'Sprint', locked: true },
            { key: 'advocates', header: 'Advocates Onboarded' },
            { key: 'referrals', header: 'Referrals Generated' },
            { key: 'trials', header: 'Active Trial Signups' },
            { key: 'paid', header: 'Paid Signups' }
          ],
          rows: [
            { sprint: 263, advocates: 0, referrals: 0, trials: 0, paid: 0 },
            { sprint: 264, advocates: 0, referrals: 0, trials: 0, paid: 0 },
            { sprint: 265, advocates: 0, referrals: 0, trials: 0, paid: 0 }
          ],
          total: { advocates: 0, referrals: 0, trials: 0, paid: 0 }
        }
      },
      {
        id: 16,
        title: 'Partnerships & Growth - Strategic Partner Program - Wix App',
        type: 'wixApp',
        moreDetailsUrl: 'https://docs.google.com/spreadsheets',
        data: {
          lifetime: { installs: 0, active: 0, paid: 0, rating: 0 },
          target: 0,
          current: 0,
          columns: [
            { key: 'sprint', header: 'Sprint', locked: true },
            { key: 'installs', header: 'Installs' },
            { key: 'uninstalls', header: 'Uninstalls' },
            { key: 'active', header: 'Active Installs' },
            { key: 'freeSignups', header: 'Free Signups' },
            { key: 'paid', header: 'Paid Signups' }
          ],
          rows: [
            { sprint: 263, installs: 0, uninstalls: 0, active: 0, freeSignups: 0, paid: 0 },
            { sprint: 264, installs: 0, uninstalls: 0, active: 0, freeSignups: 0, paid: 0 },
            { sprint: 265, installs: 0, uninstalls: 0, active: 0, freeSignups: 0, paid: 0 }
          ],
          total: { installs: 0, uninstalls: 0, active: 0, freeSignups: 0, paid: 0 }
        }
      },
      {
        id: 17,
        title: 'New subscriptions & paid signups',
        type: 'subscriptions',
        moreDetailsUrl: 'https://docs.google.com/spreadsheets',
        data: {
          rows: [
            { channel: 'New Subscription (Direct)', totalTarget: 0, targetAsOnDate: 0, actual: 0, percentage: 0 },
            { channel: 'New Subscription (Agency)', totalTarget: 0, targetAsOnDate: 0, actual: 0, percentage: 0 },
            { channel: 'Affiliate (paid signups)', totalTarget: 0, targetAsOnDate: 0, actual: 0, percentage: 0 },
            { channel: 'Ads', totalTarget: 0, targetAsOnDate: 0, actual: 0, percentage: 0 }
          ]
        }
      }
    ]
  });

  const updateSlideTitle = (slideId, newTitle) => {
    const scrollPosition = window.scrollY;
    setSprintData(prev => ({
      ...prev,
      slides: prev.slides.map(s => s.id === slideId ? { ...s, title: newTitle } : s)
    }));
    requestAnimationFrame(() => {
      window.scrollTo(0, scrollPosition);
    });
  };

  const updateMoreDetailsUrl = (slideId, url) => {
    setSprintData(prev => ({
      ...prev,
      slides: prev.slides.map(s => s.id === slideId ? { ...s, moreDetailsUrl: url } : s)
    }));
  };

  const updateSlideData = (slideId, path, value) => {
    setSprintData(prev => ({
      ...prev,
      slides: prev.slides.map(s => {
        const isNested = slideId > 1000;
        const actualId = isNested ? Math.floor(slideId / 1000) : slideId;
        
        if (s.id !== actualId) return s;
        const newSlide = JSON.parse(JSON.stringify(s));
        let current = newSlide.data;
        
        if (s.type === 'supportData') {
          current = slideId > 1000 ? current.liveChat : current.tickets;
        } else if (s.type === 'agencyLeads') {
          current = slideId > 2000 ? current.q3Performance : current.leadsConversion;
        } else if (s.type === 'rankings' && slideId === s.id && path[0] === 'rows') {
          current = current.positionChanges;
        }
        
        for (let i = 0; i < path.length - 1; i++) {
          current = current[path[i]];
        }
        current[path[path.length - 1]] = isNaN(value) ? value : Number(value);
        return newSlide;
      })
    }));
  };

  const addRow = (slideId) => {
    setSprintData(prev => ({
      ...prev,
      slides: prev.slides.map(s => {
        const isNested = slideId > 1000;
        const actualId = isNested ? Math.floor(slideId / 1000) : slideId;
        
        if (s.id !== actualId) return s;
        const newSlide = JSON.parse(JSON.stringify(s));
        
        let targetData = newSlide.data;
        if (s.type === 'supportData') {
          targetData = slideId > 1000 ? newSlide.data.liveChat : newSlide.data.tickets;
        } else if (s.type === 'agencyLeads') {
          targetData = slideId > 2000 ? newSlide.data.q3Performance : newSlide.data.leadsConversion;
        } else if (s.type === 'rankings' && slideId === s.id) {
          targetData = newSlide.data.positionChanges;
        }
        
        const rows = targetData.rows;
        if (rows && rows.length > 0) {
          const lastRow = rows[rows.length - 1];
          const newRow: any = {};
          targetData.columns.forEach(col => {
            if (col.key === 'sprint' || col.key === 'metrics' || col.key === 'quarter') {
              if (col.key === 'sprint' && typeof lastRow.sprint === 'number') {
                newRow[col.key] = lastRow.sprint + 1;
              } else {
                newRow[col.key] = '';
              }
            } else {
              newRow[col.key] = typeof lastRow[col.key] === 'number' ? 0 : '';
            }
          });
          rows.push(newRow);
          
          // Maintain sprint limit
          if (targetData.columns.some(c => c.key === 'sprint')) {
            targetData.rows = maintainSprintLimit(rows);
          }
        }
        return newSlide;
      })
    }));
  };

  const removeRow = (slideId, rowIndex) => {
    setSprintData(prev => ({
      ...prev,
      slides: prev.slides.map(s => {
        const isNested = slideId > 1000;
        const actualId = isNested ? Math.floor(slideId / 1000) : slideId;
        
        if (s.id !== actualId) return s;
        const newSlide = JSON.parse(JSON.stringify(s));
        
        let targetData = newSlide.data;
        if (s.type === 'supportData') {
          targetData = slideId > 1000 ? newSlide.data.liveChat : newSlide.data.tickets;
        } else if (s.type === 'agencyLeads') {
          targetData = slideId > 2000 ? newSlide.data.q3Performance : newSlide.data.leadsConversion;
        } else if (s.type === 'rankings' && slideId === s.id) {
          targetData = newSlide.data.positionChanges;
        }
        
        if (targetData.rows && targetData.rows.length > 1) {
          targetData.rows.splice(rowIndex, 1);
        } else {
          alert('Cannot delete the last row. At least one row is required.');
        }
        return newSlide;
      })
    }));
  };

  const addColumn = (slideId) => {
    setSprintData(prev => ({
      ...prev,
      slides: prev.slides.map(s => {
        const isNested = slideId > 1000;
        const actualId = isNested ? Math.floor(slideId / 1000) : slideId;
        
        if (s.id !== actualId) return s;
        const newSlide = JSON.parse(JSON.stringify(s));
        
        let targetData = newSlide.data;
        if (s.type === 'supportData') {
          targetData = slideId > 1000 ? newSlide.data.liveChat : newSlide.data.tickets;
        } else if (s.type === 'agencyLeads') {
          targetData = slideId > 2000 ? newSlide.data.q3Performance : newSlide.data.leadsConversion;
        } else if (s.type === 'rankings' && slideId === s.id) {
          targetData = newSlide.data.positionChanges;
        }
        
        const newColKey = `col${Date.now()}`;
        targetData.columns.push({ key: newColKey, header: 'New Column' });
        targetData.rows.forEach(row => {
          row[newColKey] = 0;
        });
        return newSlide;
      })
    }));
  };

  const removeColumn = (slideId, colKey) => {
    setSprintData(prev => ({
      ...prev,
      slides: prev.slides.map(s => {
        const isNested = slideId > 1000;
        const actualId = isNested ? Math.floor(slideId / 1000) : slideId;
        
        if (s.id !== actualId) return s;
        const newSlide = JSON.parse(JSON.stringify(s));
        
        let targetData = newSlide.data;
        if (s.type === 'supportData') {
          targetData = slideId > 1000 ? newSlide.data.liveChat : newSlide.data.tickets;
        } else if (s.type === 'agencyLeads') {
          targetData = slideId > 2000 ? newSlide.data.q3Performance : newSlide.data.leadsConversion;
        } else if (s.type === 'rankings' && slideId === s.id) {
          targetData = newSlide.data.positionChanges;
        }
        
        const nonLockedColumns = targetData.columns.filter(c => !c.locked);
        if (nonLockedColumns.length <= 1) {
          alert('Cannot delete the last non-locked column. At least one data column is required.');
          return s;
        }
        
        targetData.columns = targetData.columns.filter(c => c.key !== colKey);
        targetData.rows.forEach(row => {
          delete row[colKey];
        });
        return newSlide;
      })
    }));
  };

  const updateColumnHeader = (slideId, colKey, newHeader) => {
    const scrollPosition = window.scrollY;
    setSprintData(prev => ({
      ...prev,
      slides: prev.slides.map(s => {
        const isNested = slideId > 1000;
        const actualId = isNested ? Math.floor(slideId / 1000) : slideId;
        
        if (s.id !== actualId) return s;
        const newSlide = JSON.parse(JSON.stringify(s));
        
        let targetData = newSlide.data;
        if (s.type === 'supportData') {
          targetData = slideId > 1000 ? newSlide.data.liveChat : newSlide.data.tickets;
        } else if (s.type === 'agencyLeads') {
          targetData = slideId > 2000 ? newSlide.data.q3Performance : newSlide.data.leadsConversion;
        } else if (s.type === 'rankings' && slideId === s.id) {
          targetData = newSlide.data.positionChanges;
        }
        
        const col = targetData.columns.find(c => c.key === colKey);
        if (col) col.header = newHeader;
        return newSlide;
      })
    }));
    requestAnimationFrame(() => {
      window.scrollTo(0, scrollPosition);
    });
  };

  const exportToPDF = () => {
    if (isEditMode) {
      alert('Please save your changes (exit Edit mode) before exporting to PDF');
      return;
    }
    
    // Add a small delay to ensure any state changes are rendered
    setTimeout(() => {
      try {
        window.print();
      } catch (error) {
        console.error('Print failed:', error);
        alert('Unable to open print dialog. Please try using Ctrl+P (Windows) or Cmd+P (Mac) instead.');
      }
    }, 100);
  };

  const enterPresentMode = () => {
    setIsPresentMode(true);
    setCurrentSlide(1);
    const elem = document.documentElement;
    if (elem.requestFullscreen) {
      elem.requestFullscreen().catch(err => {
        console.log('Fullscreen request failed:', err);
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
      if (e.key === 'ArrowRight' || e.key === 'ArrowDown' || e.key === ' ') {
        e.preventDefault();
        goToNextSlide();
      } else if (e.key === 'ArrowLeft' || e.key === 'ArrowUp') {
        e.preventDefault();
        goToPreviousSlide();
      } else if (e.key === 'Escape') {
        e.preventDefault();
        exitPresentMode();
      } else if (e.key === 'Home') {
        e.preventDefault();
        setCurrentSlide(1);
      } else if (e.key === 'End') {
        e.preventDefault();
        setCurrentSlide(sprintData.slides.length);
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [isPresentMode, currentSlide, sprintData.slides.length]);

  const EditableTable = ({ slide }) => {
    const data = slide.data;
    if (!data || !data.rows || !data.columns) return null;
    const lastIdx = data.rows.length - 1;

    return (
      <div>
        {isEditMode && (
          <div style={{ marginBottom: '12px', display: 'flex', gap: '8px', flexWrap: 'wrap' }}>
            <button onClick={() => addRow(slide.id)} style={{ padding: '8px 16px', fontSize: '13px', fontWeight: '600', backgroundColor: '#2DAD70', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer' }}>+ Add Row</button>
            <button onClick={() => addColumn(slide.id)} style={{ padding: '8px 16px', fontSize: '13px', fontWeight: '600', backgroundColor: '#1863DC', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer' }}>+ Add Column</button>
          </div>
        )}
        <div style={{ overflowX: 'auto' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '14px', fontFamily: 'Inter, -apple-system, BlinkMacSystemFont, sans-serif' }}>
            <thead>
              <tr style={{ backgroundColor: '#F8FAFB', borderBottom: '2px solid #1863DC' }}>
                {data.columns.map((col) => (
                  <th key={col.key} style={{ padding: '12px 16px', textAlign: 'left', fontWeight: '600', color: '#212121' }}>
                    {isEditMode && !col.locked ? (
                      <div style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                        <input
                          key={`header-${slide.id}-${col.key}`}
                          defaultValue={col.header}
                          onBlur={(e) => updateColumnHeader(slide.id, col.key, e.target.value)}
                          style={{ padding: '6px 8px', border: '1px solid #DBDFE4', borderRadius: '4px', fontWeight: '600', flex: 1, fontSize: '14px' }}
                        />
                        <button onClick={() => removeColumn(slide.id, col.key)} style={{ padding: '4px 8px', fontSize: '12px', backgroundColor: '#DC2143', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer' }}>×</button>
                      </div>
                    ) : col.header}
                  </th>
                ))}
                {isEditMode && <th style={{ padding: '12px 16px', width: '80px' }}>Actions</th>}
              </tr>
            </thead>
            <tbody>
              {data.rows.map((row, rowIdx) => (
                <tr key={rowIdx} style={{ backgroundColor: rowIdx % 2 ? '#F8FAFB' : '#fff', borderBottom: '1px solid #EAEEF2', borderTop: rowIdx === lastIdx ? '3px solid #1863DC' : 'none' }}>
                  {data.columns.map((col, colIdx) => {
                    const currentValue = row[col.key];
                    const isLatestRow = rowIdx === lastIdx;
                    const isFirstCol = colIdx === 0;
                    const isLastCol = colIdx === data.columns.length - 1;
                    
                    return (
                      <td key={col.key} style={{ padding: '12px 16px', fontWeight: isLatestRow ? '600' : '400', color: '#212121', borderLeft: isFirstCol && isLatestRow ? '3px solid #1863DC' : 'none', borderRight: isLastCol && isLatestRow && !isEditMode ? '3px solid #1863DC' : 'none' }}>
                        {isEditMode ? (
                          <input
                            key={`${slide.id}-${rowIdx}-${col.key}`}
                            type={typeof currentValue === 'number' ? 'number' : 'text'}
                            defaultValue={currentValue ?? ''}
                            onBlur={(e) => updateSlideData(slide.id, ['rows', rowIdx, col.key], e.target.value)}
                            style={{ width: '100%', padding: '6px 8px', border: '1px solid #DBDFE4', borderRadius: '4px', fontSize: '14px' }}
                          />
                        ) : (
                          col.key === 'sprint' ? `Sprint ${currentValue}` : currentValue
                        )}
                      </td>
                    );
                  })}
                  {isEditMode && (
                    <td style={{ padding: '12px 16px', borderRight: rowIdx === lastIdx ? '3px solid #1863DC' : 'none' }}>
                      <button onClick={() => removeRow(slide.id, rowIdx)} style={{ padding: '6px 12px', fontSize: '12px', fontWeight: '500', backgroundColor: '#DC2143', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer' }}>Delete</button>
                    </td>
                  )}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    );
  };

  const Slide = ({ slide }) => (
    <div className="slide-section" style={{ backgroundColor: '#fff', padding: '32px', borderRadius: '8px', boxShadow: '0 1px 3px rgba(0,0,0,0.08)', marginBottom: '24px', pageBreakAfter: 'always' }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '20px', borderBottom: '2px solid #EAEEF2', paddingBottom: '12px' }}>
        <h2 style={{ fontSize: '20px', fontWeight: '600', color: '#212121', margin: 0, fontFamily: 'Inter, -apple-system, BlinkMacSystemFont, sans-serif', flex: 1 }}>
          {isEditMode ? (
            <input
              key={`title-${slide.id}`}
              defaultValue={slide.title}
              onBlur={(e) => updateSlideTitle(slide.id, e.target.value)}
              style={{ width: '100%', fontSize: '20px', fontWeight: '600', padding: '8px', border: '2px solid #1863DC', borderRadius: '4px' }}
            />
          ) : slide.title}
        </h2>
        <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
          {!isEditMode && slide.moreDetailsUrl && (
            <a 
              href={slide.moreDetailsUrl} 
              target="_blank" 
              rel="noopener noreferrer"
              style={{ padding: '8px 16px', fontSize: '13px', fontWeight: '600', backgroundColor: '#EBF3FD', color: '#1863DC', textDecoration: 'none', borderRadius: '20px', whiteSpace: 'nowrap', border: '1px solid #1863DC' }}
            >
              📊 More Details
            </a>
          )}
          {isEditMode && (
            <>
              <button
                onClick={() => openSlideImport(slide.id)}
                style={{ padding: '8px 16px', fontSize: '13px', fontWeight: '600', backgroundColor: '#2DAD70', color: '#fff', border: 'none', borderRadius: '20px', cursor: 'pointer', whiteSpace: 'nowrap' }}
              >
                🔄 Update Slide
              </button>
              <input
                key={`url-${slide.id}`}
                type="url"
                placeholder="Google Sheet URL (required for updates)"
                defaultValue={slide.moreDetailsUrl || ''}
                onBlur={(e) => updateMoreDetailsUrl(slide.id, e.target.value)}
                style={{ width: '280px', padding: '8px', fontSize: '13px', border: '1px solid #DBDFE4', borderRadius: '4px' }}
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
      case 'rankings':
        return (
          <div>
            <div style={{ marginBottom: '24px' }}>
              <h3 style={{ fontSize: '14px', fontWeight: '600', marginBottom: '12px', color: '#5A6872', textTransform: 'uppercase', letterSpacing: '0.5px' }}>TOTAL #1 RANKINGS</h3>
              <div style={{ fontSize: '32px', fontWeight: '700', color: '#1863DC' }}>
                {isEditMode ? <input key={`total-${slide.id}`} type="number" defaultValue={slide.data.total} onBlur={(e) => updateSlideData(slide.id, ['total'], e.target.value)} style={{ fontSize: '32px', padding: '8px', width: '120px', border: '2px solid #1863DC', borderRadius: '4px' }} /> : slide.data.total}
              </div>
            </div>
            <div style={{ marginBottom: '24px' }}>
              <h3 style={{ fontSize: '14px', fontWeight: '600', marginBottom: '12px', color: '#5A6872', textTransform: 'uppercase', letterSpacing: '0.5px' }}>#1 Rankings by Region</h3>
              {slide.data.byRegion.map((item, idx) => (
                <div key={idx} style={{ marginBottom: '12px', padding: '12px 16px', backgroundColor: '#F8FAFB', borderRadius: '6px', border: '1px solid #EAEEF2' }}>
                  <strong style={{ color: '#212121', fontSize: '14px' }}>{item.region} ({item.count})</strong><br />
                  <span style={{ fontSize: '12px', color: '#5A6872' }}>{item.keywords}</span>
                </div>
              ))}
            </div>
            <div style={{ marginBottom: '24px' }}>
              <h3 style={{ fontSize: '14px', fontWeight: '600', marginBottom: '12px', color: '#5A6872', textTransform: 'uppercase', letterSpacing: '0.5px' }}>Position Changes</h3>
              <EditableTable slide={{ ...slide, id: slide.id, data: slide.data.positionChanges }} />
            </div>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '20px' }}>
              <div>
                <h3 style={{ fontSize: '14px', fontWeight: '600', marginBottom: '12px', color: '#2DAD70' }}>↑ Improved</h3>
                {slide.data.improved.map((item, idx) => (
                  <div key={idx} style={{ marginBottom: '12px', padding: '12px 16px', backgroundColor: '#E6F7ED', borderRadius: '6px', border: '1px solid #2DAD70' }}>
                    <strong style={{ color: '#212121', fontSize: '14px' }}>{item.region} ({item.count})</strong><br />
                    <span style={{ fontSize: '12px', color: '#5A6872' }}>{item.keywords}</span>
                  </div>
                ))}
              </div>
              <div>
                <h3 style={{ fontSize: '14px', fontWeight: '600', marginBottom: '12px', color: '#DC2143' }}>↓ Declined</h3>
                {slide.data.declined.map((item, idx) => (
                  <div key={idx} style={{ marginBottom: '12px', padding: '12px 16px', backgroundColor: '#FEE9E9', borderRadius: '6px', border: '1px solid #DC2143' }}>
                    <strong style={{ color: '#212121', fontSize: '14px' }}>{item.region} ({item.count})</strong><br />
                    <span style={{ fontSize: '12px', color: '#5A6872' }}>{item.keywords}</span>
                  </div>
                ))}
              </div>
            </div>
          </div>
        );

      case 'table':
        return <EditableTable slide={slide} />;

      case 'pluginRanking':
        return (
          <div>
            <div style={{ marginBottom: '20px', padding: '16px', backgroundColor: '#EBF3FD', borderRadius: '6px', border: '1px solid #1863DC' }}>
              <strong style={{ color: '#212121', fontSize: '14px' }}>QUARTER TARGET: </strong>
              {isEditMode ? <input type="number" value={slide.data.quarterTarget} onChange={(e) => updateSlideData(slide.id, ['quarterTarget'], e.target.value)} style={{ padding: '6px 8px', width: '80px', border: '1px solid #1863DC', borderRadius: '4px' }} /> : <span style={{ fontWeight: '600', color: '#1863DC', fontSize: '16px' }}>{slide.data.quarterTarget}</span>}
            </div>
            <EditableTable slide={slide} />
          </div>
        );

      case 'supportData':
        return (
          <div>
            <div style={{ marginBottom: '32px' }}>
              <h3 style={{ fontSize: '16px', fontWeight: '600', marginBottom: '16px', color: '#212121' }}>Tickets</h3>
              <EditableTable slide={{ ...slide, id: slide.id, data: slide.data.tickets }} />
            </div>
            <div>
              <h3 style={{ fontSize: '16px', fontWeight: '600', marginBottom: '16px', color: '#212121' }}>Live Chat</h3>
              <EditableTable slide={{ ...slide, id: slide.id + 1000, data: slide.data.liveChat }} />
            </div>
          </div>
        );

      case 'agencyLeads':
        return (
          <div>
            <div style={{ marginBottom: '32px' }}>
              <EditableTable slide={{ ...slide, id: slide.id, data: slide.data.leadsConversion }} />
            </div>
            <div>
              <EditableTable slide={{ ...slide, id: slide.id + 2000, data: slide.data.q3Performance }} />
            </div>
          </div>
        );

      case 'textarea':
        return (
          <div style={{ padding: '20px', backgroundColor: '#E6F7ED', borderRadius: '8px', border: '1px solid #2DAD70', fontSize: '14px', whiteSpace: 'pre-line', color: '#212121', lineHeight: '1.8', minHeight: '200px' }}>
            {isEditMode ? <textarea key={`textarea-${slide.id}`} defaultValue={slide.data.text} onBlur={(e) => updateSlideData(slide.id, ['text'], e.target.value)} style={{ width: '100%', height: '350px', padding: '12px', fontSize: '14px', border: '1px solid #2DAD70', borderRadius: '6px', fontFamily: 'Inter, sans-serif' }} /> : (slide.data.text || 'No content yet. Click Edit to add observations.')}
          </div>
        );

      case 'quarterStats':
        return (
          <div>
            <div style={{ marginBottom: '20px' }}>
              <h3 style={{ fontSize: '16px', fontWeight: '600', marginBottom: '16px', color: '#212121' }}>Quarter Stats</h3>
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(5, 1fr)', gap: '16px' }}>
                {Object.entries(slide.data.quarterStats).map(([key, value]) => (
                  <div key={key} style={{ padding: '20px', backgroundColor: '#EBF3FD', borderRadius: '8px', textAlign: 'center', border: '1px solid #4682E1' }}>
                    <div style={{ fontSize: '36px', fontWeight: '700', color: '#1863DC', marginBottom: '8px' }}>
                      {isEditMode ? <input type="number" step="0.01" value={value as number} onChange={(e) => updateSlideData(slide.id, ['quarterStats', key], e.target.value)} style={{ width: '100px', fontSize: '36px', padding: '4px', textAlign: 'center', border: '2px solid #1863DC', borderRadius: '4px' }} /> : <span>{String(value)}</span>}
                    </div>
                    <div style={{ fontSize: '11px', color: '#5A6872', textTransform: 'uppercase', fontWeight: '600', letterSpacing: '0.5px' }}>{key.replace(/([A-Z])/g, ' $1').replace(/^./, str => str.toUpperCase())}</div>
                  </div>
                ))}
              </div>
            </div>
            <EditableTable slide={slide} />
          </div>
        );

      case 'withTarget':
        return (
          <div>
            <div style={{ marginBottom: '20px', padding: '16px', backgroundColor: '#EBF3FD', borderRadius: '6px', border: '1px solid #1863DC' }}>
              <strong style={{ color: '#212121', fontSize: '14px' }}>SPRINT TARGET: </strong>
              {isEditMode ? <input type="number" value={slide.data.target} onChange={(e) => updateSlideData(slide.id, ['target'], e.target.value)} style={{ padding: '6px 8px', width: '80px', border: '1px solid #1863DC', borderRadius: '4px' }} /> : <span style={{ fontWeight: '600', color: '#1863DC', fontSize: '16px' }}>{slide.data.target}</span>}
            </div>
            <EditableTable slide={slide} />
            {slide.data.total && (
              <div style={{ marginTop: '20px' }}>
                <h3 style={{ fontSize: '16px', fontWeight: '600', marginBottom: '16px', color: '#212121' }}>QTD Stats</h3>
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: '16px' }}>
                  {Object.entries(slide.data.total).map(([key, value]) => (
                    <div key={key} style={{ padding: '20px', backgroundColor: '#EBF3FD', borderRadius: '8px', textAlign: 'center', border: '1px solid #4682E1' }}>
                      <div style={{ fontSize: '36px', fontWeight: '700', color: '#1863DC', marginBottom: '8px' }}>
                        {isEditMode ? <input type="number" value={value as number} onChange={(e) => updateSlideData(slide.id, ['total', key], e.target.value)} style={{ width: '100px', fontSize: '36px', padding: '4px', textAlign: 'center', border: '2px solid #1863DC', borderRadius: '4px' }} /> : <span>{String(value)}</span>}
                      </div>
                      <div style={{ fontSize: '11px', color: '#5A6872', textTransform: 'uppercase', fontWeight: '600', letterSpacing: '0.5px' }}>{key}</div>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
        );

      case 'referral':
        return (
          <div>
            {slide.data.target !== undefined && (
              <div style={{ marginBottom: '20px', padding: '16px', backgroundColor: '#EBF3FD', borderRadius: '6px', border: '1px solid #1863DC' }}>
                <strong style={{ color: '#212121', fontSize: '14px' }}>TARGET VS CURRENT - PAID PLANS: </strong>
                {isEditMode ? (
                  <>
                    <input type="number" value={slide.data.target} onChange={(e) => updateSlideData(slide.id, ['target'], e.target.value)} style={{ padding: '6px 8px', width: '80px', border: '1px solid #1863DC', borderRadius: '4px', marginRight: '8px' }} />
                    <span style={{ fontWeight: '600', color: '#1863DC', fontSize: '16px' }}>Vs</span>
                    <input type="number" value={slide.data.current} onChange={(e) => updateSlideData(slide.id, ['current'], e.target.value)} style={{ padding: '6px 8px', width: '80px', border: '1px solid #1863DC', borderRadius: '4px', marginLeft: '8px' }} />
                  </>
                ) : (
                  <span style={{ fontWeight: '600', color: '#1863DC', fontSize: '16px' }}>{slide.data.target} Vs {slide.data.current}</span>
                )}
              </div>
            )}
            <EditableTable slide={slide} />
            <div style={{ marginTop: '20px' }}>
              <h3 style={{ fontSize: '16px', fontWeight: '600', marginBottom: '16px', color: '#212121' }}>App Lifetime Stats</h3>
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(2, 1fr)', gap: '16px' }}>
                {Object.entries(slide.data.lifetime).map(([key, value]) => (
                  <div key={key} style={{ padding: '20px', backgroundColor: '#EBF3FD', borderRadius: '8px', textAlign: 'center', border: '1px solid #4682E1' }}>
                    <div style={{ fontSize: '36px', fontWeight: '700', color: '#1863DC', marginBottom: '8px' }}>
                      {isEditMode ? <input type="number" value={value as number} onChange={(e) => updateSlideData(slide.id, ['lifetime', key], e.target.value)} style={{ width: '100px', fontSize: '36px', padding: '4px', textAlign: 'center', border: '2px solid #1863DC', borderRadius: '4px' }} /> : <span>{String(value)}</span>}
                    </div>
                    <div style={{ fontSize: '11px', color: '#5A6872', textTransform: 'uppercase', fontWeight: '600', letterSpacing: '0.5px' }}>{key === 'paid' ? 'Paid Signups' : key}</div>
                  </div>
                ))}
              </div>
            </div>
            {slide.data.total && (
              <div style={{ marginTop: '20px' }}>
                <h3 style={{ fontSize: '16px', fontWeight: '600', marginBottom: '16px', color: '#212121' }}>QTD Stats</h3>
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: '16px' }}>
                  {Object.entries(slide.data.total).map(([key, value]) => (
                    <div key={key} style={{ padding: '20px', backgroundColor: '#EBF3FD', borderRadius: '8px', textAlign: 'center', border: '1px solid #4682E1' }}>
                      <div style={{ fontSize: '36px', fontWeight: '700', color: '#1863DC', marginBottom: '8px' }}>
                        {isEditMode ? <input type="number" value={value as number} onChange={(e) => updateSlideData(slide.id, ['total', key], e.target.value)} style={{ width: '100px', fontSize: '36px', padding: '4px', textAlign: 'center', border: '2px solid #1863DC', borderRadius: '4px' }} /> : <span>{String(value)}</span>}
                      </div>
                      <div style={{ fontSize: '11px', color: '#5A6872', textTransform: 'uppercase', fontWeight: '600', letterSpacing: '0.5px' }}>{key}</div>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
        );

      case 'wixApp':
        return (
          <div>
            {slide.data.target !== undefined && (
              <div style={{ marginBottom: '20px', padding: '16px', backgroundColor: '#EBF3FD', borderRadius: '6px', border: '1px solid #1863DC' }}>
                <strong style={{ color: '#212121', fontSize: '14px' }}>TARGET VS CURRENT - PAID PLANS: </strong>
                {isEditMode ? (
                  <>
                    <input type="number" value={slide.data.target} onChange={(e) => updateSlideData(slide.id, ['target'], e.target.value)} style={{ padding: '6px 8px', width: '80px', border: '1px solid #1863DC', borderRadius: '4px', marginRight: '8px' }} />
                    <span style={{ fontWeight: '600', color: '#1863DC', fontSize: '16px' }}>Vs</span>
                    <input type="number" value={slide.data.current} onChange={(e) => updateSlideData(slide.id, ['current'], e.target.value)} style={{ padding: '6px 8px', width: '80px', border: '1px solid #1863DC', borderRadius: '4px', marginLeft: '8px' }} />
                  </>
                ) : (
                  <span style={{ fontWeight: '600', color: '#1863DC', fontSize: '16px' }}>{slide.data.target} Vs {slide.data.current}</span>
                )}
              </div>
            )}
            <EditableTable slide={slide} />
            <div style={{ marginTop: '20px' }}>
              <h3 style={{ fontSize: '16px', fontWeight: '600', marginBottom: '16px', color: '#212121' }}>App Lifetime Stats</h3>
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: '16px' }}>
                {Object.entries(slide.data.lifetime).map(([key, value]) => (
                  <div key={key} style={{ padding: '20px', backgroundColor: '#EBF3FD', borderRadius: '8px', textAlign: 'center', border: '1px solid #4682E1' }}>
                    <div style={{ fontSize: '36px', fontWeight: '700', color: '#1863DC', marginBottom: '8px' }}>
                      {isEditMode ? <input type="number" step={key === 'rating' ? '0.01' : '1'} value={value as number} onChange={(e) => updateSlideData(slide.id, ['lifetime', key], e.target.value)} style={{ width: '100px', fontSize: '36px', padding: '4px', textAlign: 'center', border: '2px solid #1863DC', borderRadius: '4px' }} /> : <span>{String(value)}</span>}
                    </div>
                    <div style={{ fontSize: '11px', color: '#5A6872', textTransform: 'uppercase', fontWeight: '600', letterSpacing: '0.5px' }}>{key}</div>
                  </div>
                ))}
              </div>
            </div>
          </div>
        );

      case 'subscriptions':
        return (
          <div>
            {isEditMode && (
              <div style={{ marginBottom: '16px' }}>
                <button onClick={() => {
                  const newRow = { channel: 'New Channel', totalTarget: 0, targetAsOnDate: 0, actual: 0, percentage: 0 };
                  updateSlideData(slide.id, ['rows'], [...slide.data.rows, newRow]);
                }} style={{ padding: '8px 16px', fontSize: '13px', fontWeight: '600', backgroundColor: '#2DAD70', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer' }}>+ Add Row</button>
              </div>
            )}
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '14px', fontFamily: 'Inter, sans-serif' }}>
              <thead>
                <tr style={{ backgroundColor: '#F8FAFB', borderBottom: '2px solid #1863DC' }}>
                  <th style={{ padding: '12px 16px', textAlign: 'left', fontWeight: '600', color: '#212121' }}>Channel</th>
                  <th style={{ padding: '12px 16px', textAlign: 'left', fontWeight: '600', color: '#212121' }}>Total Target</th>
                  <th style={{ padding: '12px 16px', textAlign: 'left', fontWeight: '600', color: '#212121' }}>Target as on Date</th>
                  <th style={{ padding: '12px 16px', textAlign: 'left', fontWeight: '600', color: '#212121' }}>Actual</th>
                  <th style={{ padding: '12px 16px', textAlign: 'left', fontWeight: '600', color: '#212121' }}>% (Target as on Date)</th>
                  {isEditMode && <th style={{ padding: '12px 16px' }}>Actions</th>}
                </tr>
              </thead>
              <tbody>
                {slide.data.rows.map((row, idx) => (
                  <tr key={idx} style={{ backgroundColor: idx % 2 ? '#F8FAFB' : '#fff', borderBottom: '1px solid #EAEEF2' }}>
                    <td style={{ padding: '12px 16px', fontWeight: '600', color: '#212121' }}>
                      {isEditMode ? <input key={`channel-${idx}`} defaultValue={row.channel} onBlur={(e) => updateSlideData(slide.id, ['rows', idx, 'channel'], e.target.value)} style={{ width: '100%', padding: '6px 8px', border: '1px solid #DBDFE4', borderRadius: '4px' }} /> : row.channel}
                    </td>
                    <td style={{ padding: '12px 16px' }}>
                      {isEditMode ? <input key={`target-${idx}`} type="number" defaultValue={row.totalTarget ?? ''} onBlur={(e) => updateSlideData(slide.id, ['rows', idx, 'totalTarget'], e.target.value)} style={{ width: '100px', padding: '6px 8px', border: '1px solid #DBDFE4', borderRadius: '4px' }} /> : (row.totalTarget || '-')}
                    </td>
                    <td style={{ padding: '12px 16px' }}>
                      {isEditMode ? <input key={`targetdate-${idx}`} type="number" defaultValue={row.targetAsOnDate ?? ''} onBlur={(e) => updateSlideData(slide.id, ['rows', idx, 'targetAsOnDate'], e.target.value)} style={{ width: '100px', padding: '6px 8px', border: '1px solid #DBDFE4', borderRadius: '4px' }} /> : (row.targetAsOnDate || '-')}
                    </td>
                    <td style={{ padding: '12px 16px' }}>
                      {isEditMode ? <input key={`actual-${idx}`} type="number" defaultValue={row.actual ?? ''} onBlur={(e) => updateSlideData(slide.id, ['rows', idx, 'actual'], e.target.value)} style={{ width: '100px', padding: '6px 8px', border: '1px solid #DBDFE4', borderRadius: '4px' }} /> : (row.actual || '-')}
                    </td>
                    <td style={{ padding: '12px 16px' }}>
                      {isEditMode ? <input key={`percentage-${idx}`} type="number" defaultValue={row.percentage ?? ''} onBlur={(e) => updateSlideData(slide.id, ['rows', idx, 'percentage'], e.target.value)} style={{ width: '80px', padding: '6px 8px', border: '1px solid #DBDFE4', borderRadius: '4px' }} /> : (row.percentage !== null && row.percentage !== 0 ? `${row.percentage}%` : '-')}
                    </td>
                    {isEditMode && (
                      <td style={{ padding: '12px 16px' }}>
                        <button onClick={() => {
                          const newRows = slide.data.rows.filter((_, i) => i !== idx);
                          updateSlideData(slide.id, ['rows'], newRows);
                        }} style={{ padding: '6px 12px', fontSize: '12px', fontWeight: '500', backgroundColor: '#DC2143', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer' }}>Delete</button>
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
    <div style={{ padding: '24px', fontFamily: 'Inter, -apple-system, BlinkMacSystemFont, sans-serif', backgroundColor: '#F3F5F7', minHeight: '100vh', overflowAnchor: 'none' }}>
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
        <div className="no-print" style={{ backgroundColor: '#fff', padding: '24px', marginBottom: '24px', borderRadius: '8px', boxShadow: '0 1px 3px rgba(0,0,0,0.08)', position: 'sticky', top: 0, zIndex: 100 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '20px', flexWrap: 'wrap', gap: '12px' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '20px', flexWrap: 'wrap' }}>
              <h1 style={{ fontSize: '28px', fontWeight: '700', color: '#212121', margin: 0 }}>Sprint Dashboard</h1>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px', padding: '8px 16px', backgroundColor: '#EBF3FD', borderRadius: '6px', border: '2px solid #1863DC' }}>
                <span style={{ fontSize: '14px', fontWeight: '600', color: '#5A6872' }}>Current Sprint:</span>
                {isEditMode ? (
                  <input
                    type="number"
                    value={currentSprint}
                    onChange={(e) => setCurrentSprint(Number(e.target.value))}
                    style={{ width: '80px', padding: '4px 8px', fontSize: '16px', fontWeight: '700', color: '#1863DC', border: '1px solid #1863DC', borderRadius: '4px', textAlign: 'center' }}
                  />
                ) : (
                  <span style={{ fontSize: '20px', fontWeight: '700', color: '#1863DC' }}>{currentSprint}</span>
                )}
              </div>
            </div>
            <div style={{ display: 'flex', gap: '12px', flexWrap: 'wrap' }}>
              <button onClick={() => setIsEditMode(!isEditMode)} style={{ padding: '12px 20px', fontSize: '14px', fontWeight: '600', border: 'none', borderRadius: '6px', backgroundColor: isEditMode ? '#2DAD70' : '#1863DC', color: '#fff', cursor: 'pointer', boxShadow: '0 2px 4px rgba(0,0,0,0.1)' }}>{isEditMode ? '✓ Save' : '✎ Edit'}</button>
              <button onClick={() => setShowImportModal(true)} style={{ padding: '12px 20px', fontSize: '14px', fontWeight: '600', border: 'none', borderRadius: '6px', backgroundColor: '#2DAD70', color: '#fff', cursor: 'pointer', boxShadow: '0 2px 4px rgba(0,0,0,0.1)' }}>📊 Import Data</button>
              <button onClick={exportToPDF} style={{ padding: '12px 20px', fontSize: '14px', fontWeight: '600', border: 'none', borderRadius: '6px', backgroundColor: '#7F56D9', color: '#fff', cursor: 'pointer', boxShadow: '0 2px 4px rgba(0,0,0,0.1)' }}>📄 Export PDF</button>
              <button onClick={enterPresentMode} style={{ padding: '12px 20px', fontSize: '14px', fontWeight: '600', border: 'none', borderRadius: '6px', backgroundColor: '#363F52', color: '#fff', cursor: 'pointer', boxShadow: '0 2px 4px rgba(0,0,0,0.1)' }}>🎬 Present</button>
            </div>
          </div>

          <div style={{ backgroundColor: '#FEF9E6', border: '1px solid #FFB800', borderRadius: '6px', padding: '16px', fontSize: '13px', color: '#5A4A00', marginBottom: '12px' }}>
            <strong>💡 Edit Mode:</strong> Edit slide titles, add/remove rows & columns, edit all values, rename headers. Latest sprint row is bold with a blue border.
            <br />
            <strong>📊 Import Data:</strong> Upload CSV or paste data to auto-fill. Maximum 5 sprints will be kept (oldest removed automatically).
          </div>
        </div>
      )}

      {showImportModal && (
        <div style={{ position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, backgroundColor: 'rgba(0,0,0,0.5)', zIndex: 2000, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
          <div style={{ backgroundColor: '#fff', borderRadius: '12px', padding: '32px', maxWidth: '600px', width: '90%', boxShadow: '0 8px 32px rgba(0,0,0,0.3)' }}>
            <h2 style={{ fontSize: '24px', fontWeight: '700', color: '#212121', marginBottom: '20px' }}>
              Update Slide Data from Google Sheet
            </h2>
            <p style={{ fontSize: '14px', color: '#5A6872', marginBottom: '20px' }}>
              This will pull data from the Google Sheet URL attached to this slide. The system will intelligently identify and import only the relevant data for this specific slide, maintaining the last 5 sprints automatically.
            </p>
            
            <div style={{ marginBottom: '20px', padding: '16px', backgroundColor: '#E6F7ED', borderRadius: '6px', border: '1px solid #2DAD70' }}>
              <div style={{ fontSize: '13px', color: '#1D7A47', marginBottom: '8px' }}>
                <strong>✓ Smart Import:</strong>
              </div>
              <ul style={{ margin: 0, paddingLeft: '20px', fontSize: '13px', color: '#1D7A47', lineHeight: '1.6' }}>
                <li>Automatically identifies correct data from your sheet</li>
                <li>Ignores additional/irrelevant data in the sheet</li>
                <li>Maintains 5-sprint limit (removes oldest data)</li>
                <li>Only updates this specific slide</li>
              </ul>
            </div>

            <div style={{ marginBottom: '20px' }}>
              <label style={{ display: 'block', marginBottom: '8px', fontSize: '14px', fontWeight: '600', color: '#212121' }}>Or Upload CSV Manually</label>
              <input 
                type="file" 
                accept=".csv"
                onChange={handleFileUpload}
                style={{ width: '100%', padding: '12px', border: '2px dashed #1863DC', borderRadius: '6px', cursor: 'pointer', fontSize: '13px' }}
              />
            </div>

            <div style={{ marginBottom: '20px', padding: '16px', backgroundColor: '#EBF3FD', borderRadius: '6px', fontSize: '13px', color: '#134FB0' }}>
              <strong>💡 Best Practice:</strong> Keep your Google Sheet organized with clear column headers matching the slide data structure. The system will automatically map columns to the correct fields.
            </div>

            <div style={{ display: 'flex', gap: '12px', justifyContent: 'flex-end' }}>
              <button onClick={() => { setShowImportModal(false); setCurrentImportSlideId(null); }} style={{ padding: '12px 24px', fontSize: '14px', fontWeight: '600', border: '2px solid #DBDFE4', borderRadius: '6px', backgroundColor: '#fff', color: '#5A6872', cursor: 'pointer' }}>Cancel</button>
              <button 
                onClick={() => {
                  alert('Google Sheets API integration: In production, this would fetch data directly from the linked Google Sheet. For now, please use CSV upload.');
                  setShowImportModal(false);
                  setCurrentImportSlideId(null);
                }}
                style={{ padding: '12px 24px', fontSize: '14px', fontWeight: '600', border: 'none', borderRadius: '6px', backgroundColor: '#2DAD70', color: '#fff', cursor: 'pointer' }}
              >
                🔄 Update from Sheet
              </button>
            </div>
          </div>
        </div>
      )}

      {isPresentMode && (
        <div style={{ position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, backgroundColor: '#1C2630', zIndex: 1000, display: 'flex', flexDirection: 'column', overflow: 'auto' }}>
          <div style={{ position: 'fixed', top: '24px', right: '24px', zIndex: 1001, display: 'flex', gap: '12px', alignItems: 'center' }}>
            <div style={{ color: '#fff', fontSize: '16px', fontWeight: '500', backgroundColor: 'rgba(0,0,0,0.3)', padding: '8px 16px', borderRadius: '6px' }}>
              {currentSlide} / {sprintData.slides.length}
            </div>
            <button onClick={exitPresentMode} style={{ padding: '12px 24px', backgroundColor: '#DC2143', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: '600', fontSize: '14px' }}>✕ Exit (Esc)</button>
          </div>
          
          <div style={{ flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '80px 40px 40px 40px', position: 'relative' }}>
            <button 
              onClick={goToPreviousSlide} 
              disabled={currentSlide === 1}
              style={{ 
                position: 'absolute', 
                left: '20px', 
                top: '50%', 
                transform: 'translateY(-50%)',
                padding: '16px 20px', 
                backgroundColor: currentSlide === 1 ? '#5A6872' : '#1863DC', 
                color: '#fff', 
                border: 'none', 
                borderRadius: '8px', 
                cursor: currentSlide === 1 ? 'not-allowed' : 'pointer', 
                fontWeight: '600', 
                fontSize: '18px',
                opacity: currentSlide === 1 ? 0.5 : 1,
                zIndex: 10
              }}>
              ←
            </button>
            
            <div style={{ maxWidth: '1400px', width: '100%', backgroundColor: '#fff', borderRadius: '12px', boxShadow: '0 8px 32px rgba(0,0,0,0.3)', overflow: 'auto', maxHeight: '90vh' }}>
              {sprintData.slides[currentSlide - 1] && (
                <div style={{ padding: '48px' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '32px', borderBottom: '3px solid #1863DC', paddingBottom: '16px' }}>
                    <h2 style={{ fontSize: '28px', fontWeight: '700', color: '#212121', margin: 0 }}>
                      {sprintData.slides[currentSlide - 1].title}
                    </h2>
                    {sprintData.slides[currentSlide - 1].moreDetailsUrl && (
                      <a 
                        href={sprintData.slides[currentSlide - 1].moreDetailsUrl} 
                        target="_blank" 
                        rel="noopener noreferrer"
                        style={{ padding: '10px 20px', fontSize: '14px', fontWeight: '600', backgroundColor: '#EBF3FD', color: '#1863DC', textDecoration: 'none', borderRadius: '24px', whiteSpace: 'nowrap', border: '1px solid #1863DC' }}
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
                position: 'absolute', 
                right: '20px', 
                top: '50%', 
                transform: 'translateY(-50%)',
                padding: '16px 20px', 
                backgroundColor: currentSlide === sprintData.slides.length ? '#5A6872' : '#1863DC', 
                color: '#fff', 
                border: 'none', 
                borderRadius: '8px', 
                cursor: currentSlide === sprintData.slides.length ? 'not-allowed' : 'pointer', 
                fontWeight: '600', 
                fontSize: '18px',
                opacity: currentSlide === sprintData.slides.length ? 0.5 : 1,
                zIndex: 10
              }}>
              →
            </button>
          </div>

          <div style={{ padding: '20px', textAlign: 'center', color: '#fff', fontSize: '13px', backgroundColor: 'rgba(0,0,0,0.2)' }}>
            Use ← → arrow keys, Space, or on-screen buttons to navigate | Press Esc to exit
          </div>
        </div>
      )}

      <div style={{ maxWidth: '1400px', margin: '0 auto', display: isPresentMode ? 'none' : 'block' }}>
        {sprintData.slides.map(slide => (
          <Slide key={slide.id} slide={slide} />
        ))}
      </div>
    </div>
  );
}