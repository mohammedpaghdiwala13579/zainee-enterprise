import React, { useState, useEffect } from "react";
import { Download, Monitor, Smartphone, X, Check, Info, Share2, PlusSquare, ArrowUpRight, Laptop } from "lucide-react";
import QuotationBuilder from "./components/QuotationBuilder";

export default function App() {
  const [deferredPrompt, setDeferredPrompt] = useState<any>(null);
  const [isInstallable, setIsInstallable] = useState(false);
  const [isStandalone, setIsStandalone] = useState(false);
  const [showBanner, setShowBanner] = useState(false);
  const [showGuidance, setShowGuidance] = useState(false);
  const [platform, setPlatform] = useState<"ios" | "android" | "mac" | "windows" | "other">("other");

  useEffect(() => {
    // 1. Check if the application is already running in standalone/installed mode
    const checkStandalone = () => {
      const isStandaloneMode = 
        window.matchMedia('(display-mode: standalone)').matches || 
        (window.navigator as any).standalone === true;
      setIsStandalone(isStandaloneMode);
      
      // Only show banner if NOT standalone and NOT dismissed in this session
      const isDismissed = sessionStorage.getItem("zainee_pwa_dismissed") === "true";
      if (!isStandaloneMode && !isDismissed) {
        setShowBanner(true);
      }
    };

    checkStandalone();

    // 2. Detect device platform for customized PWA experience
    const detectPlatform = () => {
      const ua = navigator.userAgent.toLowerCase();
      if (/iphone|ipad|ipod/.test(ua)) {
        setPlatform("ios");
      } else if (/android/.test(ua)) {
        setPlatform("android");
      } else if (/macintosh|mac os x/.test(ua)) {
        setPlatform("mac");
      } else if (/windows|win32|win64/.test(ua)) {
        setPlatform("windows");
      } else {
        setPlatform("other");
      }
    };

    detectPlatform();

    // 3. Listen for the native browser install prompt event
    const handleBeforeInstallPrompt = (e: Event) => {
      // Prevent standard browser bar from displaying
      e.preventDefault();
      // Store the event so it can be triggered with a single click
      setDeferredPrompt(e);
      setIsInstallable(true);
      
      const isDismissed = sessionStorage.getItem("zainee_pwa_dismissed") === "true";
      if (!isStandalone && !isDismissed) {
        setShowBanner(true);
      }
    };

    window.addEventListener("beforeinstallprompt", handleBeforeInstallPrompt);

    // 4. Listen for successful installation
    const handleAppInstalled = () => {
      setIsStandalone(true);
      setIsInstallable(false);
      setDeferredPrompt(null);
      setShowBanner(false);
      setShowGuidance(false);
      console.log("Zainee Enterprise Terminal was successfully installed!");
    };

    window.addEventListener("appinstalled", handleAppInstalled);

    return () => {
      window.removeEventListener("beforeinstallprompt", handleBeforeInstallPrompt);
      window.removeEventListener("appinstalled", handleAppInstalled);
    };
  }, [isStandalone]);

  // Single click install trigger
  const handleInstallClick = async () => {
    if (deferredPrompt) {
      // Chrome, Edge, and Android (Chrome) support direct native prompt triggering
      deferredPrompt.prompt();
      const choiceResult = await deferredPrompt.userChoice;
      if (choiceResult.outcome === 'accepted') {
        console.log('User accepted the PWA install prompt');
        setIsStandalone(true);
        setShowBanner(false);
      } else {
        console.log('User dismissed the PWA install prompt');
      }
      setDeferredPrompt(null);
    } else {
      // Fallback/guidance for devices like iOS (Safari), macOS (Safari), or browsers without prompt events
      setShowGuidance(true);
    }
  };

  const dismissBanner = () => {
    setShowBanner(false);
    sessionStorage.setItem("zainee_pwa_dismissed", "true");
  };

  return (
    <div className="min-h-screen bg-slate-100 print:bg-white print:p-0 print:m-0 text-slate-900 font-sans flex flex-col selection:bg-blue-600 selection:text-white">
      
      {/* 1. Dynamic PWA Install Banner - Top of the page */}
      {showBanner && !isStandalone && (
        <div id="pwa-install-banner" className="no-print bg-[#0b132b] text-white py-2 px-4 sm:px-6 relative flex flex-col md:flex-row items-center justify-between gap-3 shadow-md border-b border-slate-800 z-50 animate-in slide-in-from-top duration-300">
          <div className="flex items-center gap-3">
            <div className="bg-blue-600/30 p-1.5 rounded border border-blue-500/30 flex items-center justify-center shrink-0">
              {platform === "ios" || platform === "android" ? (
                <Smartphone className="h-4 w-4 text-blue-400" />
              ) : (
                <Laptop className="h-4 w-4 text-blue-400" />
              )}
            </div>
            <div>
              <p className="text-[12px] font-semibold tracking-wide flex items-center gap-2">
                <span>COMILLA TRADERS ERP TERMINAL</span>
                <span className="bg-blue-600/30 text-blue-300 text-[9px] font-bold px-1.5 py-0.5 rounded border border-blue-500/30 uppercase tracking-wider">Fast &amp; Secure</span>
              </p>
              <p className="text-[11px] text-slate-300 font-normal">
                Install on your {platform === "ios" ? "iPhone/iPad" : platform === "mac" ? "Macbook" : platform === "windows" ? "Windows PC" : platform === "android" ? "Android Phone" : "Device"} for standalone ERP desktop access and offline capability.
              </p>
            </div>
          </div>
          
          <div className="flex items-center gap-2 shrink-0">
            <button
              id="btn-banner-install"
              onClick={handleInstallClick}
              className="bg-blue-600 hover:bg-blue-700 text-white text-xs font-semibold px-3 py-1.5 rounded flex items-center gap-1.5 shadow-xs transition-colors cursor-pointer shrink-0"
              title="Install ERP App"
            >
              <Download className="h-3.5 w-3.5" />
              <span>Install Terminal</span>
            </button>
            <button
              id="btn-banner-dismiss"
              onClick={dismissBanner}
              className="text-slate-400 hover:text-white p-1 hover:bg-white/10 rounded transition-colors cursor-pointer"
              title="Dismiss"
            >
              <X className="h-4 w-4" />
            </button>
          </div>
        </div>
      )}

      {/* 2. iOS / Safari / Manual Install Guidance Modal */}
      {showGuidance && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 z-50 animate-in fade-in duration-200">
          <div className="bg-white rounded-lg shadow-2xl border border-slate-200 max-w-md w-full p-6 relative animate-in zoom-in-95 duration-200">
            <button
              id="btn-modal-close"
              onClick={() => setShowGuidance(false)}
              className="absolute top-4 right-4 text-slate-400 hover:text-slate-600 p-1 rounded hover:bg-slate-100 transition-colors cursor-pointer"
            >
              <X className="h-5 w-5" />
            </button>

            <div className="flex items-center gap-3 mb-4">
              <div className="h-9 w-9 rounded bg-blue-50 flex items-center justify-center text-blue-600 shrink-0">
                <Info className="h-5 w-5" />
              </div>
              <div>
                <h3 className="text-sm font-bold text-slate-900">How to Install Application</h3>
                <p className="text-[11px] text-slate-500">Enable single-click ERP terminal launching</p>
              </div>
            </div>

            <div className="space-y-3 my-4">
              {/* iOS Safari Guidance */}
              {platform === "ios" ? (
                <div className="bg-slate-50 p-3.5 rounded border border-slate-200 text-xs space-y-2.5">
                  <p className="font-semibold text-slate-900 flex items-center gap-1.5">
                    <Smartphone className="h-4 w-4 text-blue-600" />
                    For iPhone &amp; iPad (Safari):
                  </p>
                  <ol className="list-decimal pl-4 space-y-1.5 text-slate-700 text-[11px]">
                    <li>
                      Tap the <strong className="text-slate-900">Share</strong> button ( <Share2 className="h-3 w-3 inline-block text-blue-600" /> icon) at the bottom or top of Safari.
                    </li>
                    <li>
                      Scroll down and select <strong className="text-slate-900">Add to Home Screen</strong> ( <PlusSquare className="h-3 w-3 inline-block text-blue-600" /> icon).
                    </li>
                    <li>
                      Tap <strong className="text-slate-900">Add</strong> in the top-right corner.
                    </li>
                  </ol>
                  <p className="text-[10px] text-slate-500 italic mt-1">Note: This is Apple's standard requirement for web applications.</p>
                </div>
              ) : platform === "mac" ? (
                /* macOS Safari Guidance */
                <div className="bg-slate-50 p-3.5 rounded border border-slate-200 text-xs space-y-2.5">
                  <p className="font-semibold text-slate-900 flex items-center gap-1.5">
                    <Laptop className="h-4 w-4 text-blue-600" />
                    For macOS (Safari):
                  </p>
                  <ol className="list-decimal pl-4 space-y-1.5 text-slate-700 text-[11px]">
                    <li>
                      Go to the top browser menu and click <strong className="text-slate-900">File</strong>.
                    </li>
                    <li>
                      Select <strong className="text-slate-900">Add to Dock...</strong> from the dropdown menu.
                    </li>
                    <li>
                      Confirm the name and click <strong className="text-slate-900">Add</strong>.
                    </li>
                  </ol>
                  <p className="text-[10px] text-slate-500 italic mt-1">This will place a Zainee Enterprise icon right in your Mac Dock!</p>
                </div>
              ) : (
                /* PC / Other Browsers manual fallback instruction */
                <div className="bg-slate-50 p-3.5 rounded border border-slate-200 text-xs space-y-2.5">
                  <p className="font-semibold text-slate-900 flex items-center gap-1.5">
                    <Monitor className="h-4 w-4 text-blue-600" />
                    Standard Browser Installation:
                  </p>
                  <ol className="list-decimal pl-4 space-y-1.5 text-slate-700 text-[11px]">
                    <li>
                      Look at the right side of your browser's address bar at the top.
                    </li>
                    <li>
                      Click the <strong className="text-slate-900">Install app</strong> icon (usually a small monitor with a down arrow, or a plus symbol).
                    </li>
                    <li>
                      If not found, open the browser menu ( <strong className="text-slate-900">three dots ⋮</strong> or <strong className="text-slate-900">hamburger menu ☰</strong> ) and click <strong className="text-slate-900">Save and share</strong> ➡️ <strong className="text-slate-900">Install page...</strong>.
                    </li>
                  </ol>
                </div>
              )}

              {/* Universal Benefits */}
              <div className="bg-slate-50 p-3 rounded border border-slate-200 text-[11px] text-slate-600 space-y-1.5">
                <p className="font-semibold text-slate-800">Benefits of Installing:</p>
                <div className="grid grid-cols-2 gap-2 text-[10px]">
                  <div className="flex items-center gap-1">
                    <Check className="h-3 w-3 text-emerald-600" /> Offline Access Ready
                  </div>
                  <div className="flex items-center gap-1">
                    <Check className="h-3 w-3 text-emerald-600" /> Zero Browser Margins
                  </div>
                  <div className="flex items-center gap-1">
                    <Check className="h-3 w-3 text-emerald-600" /> Desktop/Dock Shortcut
                  </div>
                  <div className="flex items-center gap-1">
                    <Check className="h-3 w-3 text-emerald-600" /> Fluid Window
                  </div>
                </div>
              </div>
            </div>

            <button
              id="btn-guidance-ok"
              onClick={() => setShowGuidance(false)}
              className="w-full bg-[#0b132b] hover:bg-[#152347] text-white text-xs font-semibold py-2 rounded transition-colors cursor-pointer mt-3"
            >
              Got It
            </button>
          </div>
        </div>
      )}

      {/* Main ERP Full-width Viewport */}
      <main className="flex-1 flex flex-col w-full print:p-0 print:m-0 print:bg-white print:block">
        <QuotationBuilder />
      </main>
    </div>
  );
}
