import {
  AlertCircle,
  Check,
  ImagePlus,
  Layers,
  Library,
  Loader2,
  MousePointer2,
  RefreshCw,
  Ruler,
  Trash2,
  Upload,
} from "lucide-react";
import { ChangeEvent, useEffect, useMemo, useRef, useState } from "react";
import { createPowerPointAdapter } from "./adapters/powerPointAdapter";
import { createLogoAssetFromFile, isSupportedLogoFile, isTransparentPng, roundPoint } from "./domain/image";
import { logoRepository } from "./domain/logoRepository";
import type { LogoAsset, LogoPlacement, PowerPointAdapter, PowerPointCapabilities } from "./domain/types";

type BusyState = "idle" | "importing" | "inserting" | "capturing" | "applying" | "removing";

export function App() {
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const activePlacementRef = useRef<LogoPlacement | null>(null);
  const [assets, setAssets] = useState<LogoAsset[]>([]);
  const [selectedLogoId, setSelectedLogoId] = useState<string>("");
  const [activePlacement, setActivePlacement] = useState<LogoPlacement | null>(null);
  const [adapter, setAdapter] = useState<PowerPointAdapter | null>(null);
  const [capabilities, setCapabilities] = useState<PowerPointCapabilities | null>(null);
  const [busy, setBusy] = useState<BusyState>("idle");
  const [message, setMessage] = useState<string>("先导入一个 Logo。");
  const [error, setError] = useState<string>("");

  const selectedLogo = useMemo(
    () => assets.find((asset) => asset.id === selectedLogoId) ?? assets[0] ?? null,
    [assets, selectedLogoId],
  );

  useEffect(() => {
    void bootstrap();
  }, []);

  function setCurrentPlacement(placement: LogoPlacement | null) {
    activePlacementRef.current = placement;
    setActivePlacement(placement);
  }

  async function bootstrap() {
    try {
      const activeAdapter = await createPowerPointAdapter();
      setAdapter(activeAdapter);
      const [storedAssets, storedPlacements, detectedCapabilities] = await Promise.all([
        logoRepository.listAssets(),
        logoRepository.listPlacements(),
        activeAdapter.getCapabilities(),
      ]);
      setAssets(storedAssets);
      setCapabilities(detectedCapabilities);
      setSelectedLogoId(storedAssets[0]?.id ?? "");
      setCurrentPlacement(storedPlacements[0] ?? null);
      if (storedAssets.length > 0) {
        setMessage("选择 Logo，插入当前页，调整位置后读取定位。");
      }
    } catch (caught) {
      setError(errorMessage(caught));
    }
  }

  async function handleImport(event: ChangeEvent<HTMLInputElement>) {
    const files = Array.from(event.target.files ?? []);
    event.target.value = "";
    if (files.length === 0) {
      return;
    }

    await runBusy("importing", async () => {
      const validFiles = files.filter(isSupportedLogoFile);
      if (validFiles.length !== files.length) {
        setMessage(`已跳过 ${files.length - validFiles.length} 个不支持的文件。`);
      }

      const imported = await Promise.all(validFiles.map(createLogoAssetFromFile));
      await Promise.all(imported.map((asset) => logoRepository.saveAsset(asset)));
      const nextAssets = [...imported, ...assets];
      setAssets(nextAssets);
      setSelectedLogoId(imported[0]?.id ?? selectedLogoId);
      setMessage(`已导入 ${imported.length} 个 Logo 文件。`);
    });
  }

  async function insertSelectedLogo() {
    if (!adapter) {
      setError("插件还在初始化，请稍后再试。");
      return;
    }

    if (!selectedLogo) {
      setError("请先导入或选择一个 Logo。");
      return;
    }

    await runBusy("inserting", async () => {
      const inserted = await adapter.insertLogoOnCurrentSlide(selectedLogo);
      setCurrentPlacement(inserted.placement);
      await logoRepository.savePlacement(inserted.placement);
      await refreshPlacements(inserted.placement.id);
      setMessage("Logo 已插入。请在 PowerPoint 页面中拖拽或缩放，然后读取位置。");
    });
  }

  async function captureSelectedPlacement() {
    if (!adapter) {
      setError("插件还在初始化，请稍后再试。");
      return;
    }

    if (!selectedLogo) {
      setError("读取位置前，请先选择一个 Logo。");
      return;
    }

    await runBusy("capturing", async () => {
      const placement = await adapter.readSelectedLogoPlacement(selectedLogo.id);
      const normalized = { ...placement, logoId: placement.logoId || selectedLogo.id };
      setCurrentPlacement(normalized);
      await logoRepository.savePlacement(normalized);
      await refreshPlacements(normalized.id);
      setMessage("已从当前选中的 PowerPoint Logo 读取位置。");
    });
  }

  async function applyToAllSlides() {
    if (!adapter) {
      setError("插件还在初始化，请稍后再试。");
      return;
    }

    const currentPlacement = activePlacementRef.current ?? activePlacement;
    if (!selectedLogo || !currentPlacement) {
      setError("应用前，请先选择 Logo 并读取位置。");
      return;
    }

    await runBusy("applying", async () => {
      let placement = { ...currentPlacement, logoId: currentPlacement.logoId || selectedLogo.id };
      if (capabilities?.platform === "windows") {
        try {
          const selectedPlacement = await adapter.readSelectedLogoPlacement(selectedLogo.id);
          placement = { ...selectedPlacement, logoId: selectedPlacement.logoId || selectedLogo.id };
        } catch {
          // If no PowerPoint shape is selected, use the placement currently shown in the task pane.
        }
      }
      const asset = assets.find((item) => item.id === placement.logoId) ?? selectedLogo;
      setCurrentPlacement(placement);
      await adapter.applyPlacementToAllSlides(asset, placement);
      await logoRepository.savePlacement(placement);
      await refreshPlacements(placement.id);
      setMessage("已把 Logo 位置应用到所有幻灯片。");
    });
  }

  async function updateAllPositions() {
    if (!adapter) {
      setError("插件还在初始化，请稍后再试。");
      return;
    }

    const currentPlacement = activePlacementRef.current ?? activePlacement;
    if (!currentPlacement) {
      setError("请先读取一个位置。");
      return;
    }

    await runBusy("applying", async () => {
      await adapter.updatePlacement(currentPlacement);
      await logoRepository.savePlacement(currentPlacement);
      await refreshPlacements(currentPlacement.id);
      setMessage("已用当前数值更新现有 Logo。");
    });
  }

  async function removeActivePlacement() {
    if (!adapter) {
      setError("插件还在初始化，请稍后再试。");
      return;
    }

    const currentPlacement = activePlacementRef.current ?? activePlacement;
    if (!currentPlacement) {
      setError("当前没有可移除的位置记录。");
      return;
    }

    await runBusy("removing", async () => {
      await adapter.removePlacement(currentPlacement.id);
      await logoRepository.deletePlacement(currentPlacement.id);
      setCurrentPlacement(null);
      await refreshPlacements();
      setMessage("已从演示文稿中移除 Logo。");
    });
  }

  async function deleteLogo(assetId: string) {
    await runBusy("removing", async () => {
      await logoRepository.deleteAsset(assetId);
      const nextAssets = assets.filter((asset) => asset.id !== assetId);
      setAssets(nextAssets);
      if (selectedLogoId === assetId) {
        setSelectedLogoId(nextAssets[0]?.id ?? "");
      }
      setMessage("已从本地 Logo 库删除。");
    });
  }

  async function refreshPlacements(preferredId?: string) {
    const nextPlacements = await logoRepository.listPlacements();
    if (preferredId) {
      setCurrentPlacement(nextPlacements.find((item) => item.id === preferredId) ?? null);
    }
  }

  async function runBusy(nextBusy: BusyState, action: () => Promise<void>) {
    setBusy(nextBusy);
    setError("");
    try {
      await action();
    } catch (caught) {
      setError(errorMessage(caught));
    } finally {
      setBusy("idle");
    }
  }

  function updatePlacementField(field: keyof Pick<LogoPlacement, "left" | "top" | "width" | "height">, value: string) {
    const currentPlacement = activePlacementRef.current ?? activePlacement;
    if (!currentPlacement) {
      return;
    }

    const numericValue = Number(value);
    if (!Number.isFinite(numericValue)) {
      return;
    }

    setCurrentPlacement({ ...currentPlacement, [field]: roundPoint(numericValue) });
  }

  const isBusy = busy !== "idle";

  return (
    <main className="app-shell">
      <header className="app-header">
        <div>
          <h1>Logo 添加工具</h1>
          <p>把一个 Logo 添加到当前 PowerPoint 的每一页。</p>
        </div>
      </header>

      {error ? (
        <>
          <div className="notice notice-error">
            <AlertCircle size={16} />
            <span>{error}</span>
          </div>
          <EnvironmentDiagnostics capabilities={capabilities} />
        </>
      ) : (
        <div className="notice notice-success">
          <Check size={16} />
          <span>{message}</span>
        </div>
      )}

      <section className="toolbar" aria-label="Logo 导入">
        <input
          ref={fileInputRef}
          className="visually-hidden"
          type="file"
          accept="image/png,image/jpeg,image/webp"
          multiple
          onChange={handleImport}
        />
        <button type="button" className="primary-button" onClick={() => fileInputRef.current?.click()} disabled={isBusy}>
          {busy === "importing" ? <Loader2 className="spin" size={16} /> : <Upload size={16} />}
          导入 Logo
        </button>
        <button type="button" onClick={insertSelectedLogo} disabled={isBusy || !selectedLogo}>
          {busy === "inserting" ? <Loader2 className="spin" size={16} /> : <ImagePlus size={16} />}
          插入当前页
        </button>
      </section>

      <section className="panel">
        <div className="section-title">
          <Library size={16} />
          <h2>Logo 库</h2>
        </div>
        {assets.length === 0 ? (
          <div className="empty-state">还没有 Logo。请导入 PNG、JPEG 或 WebP 文件。</div>
        ) : (
          <div className="logo-grid">
            {assets.map((asset) => (
              <button
                className={`logo-card ${asset.id === selectedLogo?.id ? "selected" : ""}`}
                key={asset.id}
                type="button"
                onClick={() => setSelectedLogoId(asset.id)}
              >
                <img src={asset.data} alt={asset.name} />
                <span title={asset.name}>{asset.name}</span>
                <small>
                  {asset.intrinsicWidth} x {asset.intrinsicHeight}
                </small>
              </button>
            ))}
          </div>
        )}
        {selectedLogo ? (
          <div className="asset-actions">
            <span>{isTransparentPng(selectedLogo) ? "透明 PNG，可直接使用" : "PNG 清理能力将在第二阶段加入"}</span>
            <button type="button" className="icon-button" onClick={() => deleteLogo(selectedLogo.id)} disabled={isBusy}>
              <Trash2 size={16} />
            </button>
          </div>
        ) : null}
      </section>

      <section className="panel">
        <div className="section-title">
          <MousePointer2 size={16} />
          <h2>读取位置</h2>
        </div>
        <p className="compact-copy">插入后，直接在 PowerPoint 页面里拖拽和缩放 Logo。</p>
        <button type="button" className="wide-button" onClick={captureSelectedPlacement} disabled={isBusy || !selectedLogo}>
          {busy === "capturing" ? <Loader2 className="spin" size={16} /> : <Ruler size={16} />}
          使用当前选中 Logo 的位置
        </button>
      </section>

      <section className="panel">
        <div className="section-title">
          <Layers size={16} />
          <h2>应用到整份 PPT</h2>
        </div>
        <PlacementEditor placement={activePlacement} onChange={updatePlacementField} />
        <div className="action-grid">
          <button type="button" className="primary-button" onClick={applyToAllSlides} disabled={isBusy || !selectedLogo || !activePlacement}>
            {busy === "applying" ? <Loader2 className="spin" size={16} /> : <Layers size={16} />}
            应用全部
          </button>
          <button type="button" onClick={updateAllPositions} disabled={isBusy || !activePlacement}>
            <RefreshCw size={16} />
            更新
          </button>
          <button type="button" onClick={removeActivePlacement} disabled={isBusy || !activePlacement}>
            <Trash2 size={16} />
            移除
          </button>
        </div>
      </section>
    </main>
  );
}

function EnvironmentDiagnostics({ capabilities }: { capabilities: PowerPointCapabilities | null }) {
  if (!capabilities) {
    return (
      <div className="diagnostic-stack">
        <CapabilityBadge capabilities={capabilities} />
      </div>
    );
  }

  return (
    <div className="diagnostic-stack">
      <CapabilityBadge capabilities={capabilities} />
      {capabilities.warnings.map((warning) => (
        <div className="notice" key={warning}>
          <AlertCircle size={16} />
          <span>{warning}</span>
        </div>
      ))}
    </div>
  );
}

function CapabilityBadge({ capabilities }: { capabilities: PowerPointCapabilities | null }) {
  if (!capabilities) {
    return <span className="capability-badge">检测中</span>;
  }

  return (
    <span className="capability-badge">
      {platformLabel(capabilities.platform)} / {hostLabel(capabilities.host)}
    </span>
  );
}

function platformLabel(platform: PowerPointCapabilities["platform"]): string {
  const labels: Record<PowerPointCapabilities["platform"], string> = {
    windows: "Windows",
    mac: "Mac",
    web: "网页版",
    unknown: "未知",
  };
  return labels[platform];
}

function hostLabel(host: PowerPointCapabilities["host"]): string {
  const labels: Record<PowerPointCapabilities["host"], string> = {
    "office-js": "Office.js",
    "native-windows": "Windows 原生",
    mock: "模拟模式",
  };
  return labels[host];
}

function PlacementEditor({
  placement,
  onChange,
}: {
  placement: LogoPlacement | null;
  onChange: (field: keyof Pick<LogoPlacement, "left" | "top" | "width" | "height">, value: string) => void;
}) {
  if (!placement) {
    return <div className="empty-state">还没有读取位置。</div>;
  }

  return (
    <div className="placement-editor">
      <label>
        左边距
        <input type="number" value={placement.left} onChange={(event) => onChange("left", event.target.value)} />
      </label>
      <label>
        上边距
        <input type="number" value={placement.top} onChange={(event) => onChange("top", event.target.value)} />
      </label>
      <label>
        宽度
        <input type="number" min="1" value={placement.width} onChange={(event) => onChange("width", event.target.value)} />
      </label>
      <label>
        高度
        <input type="number" min="1" value={placement.height} onChange={(event) => onChange("height", event.target.value)} />
      </label>
    </div>
  );
}

function errorMessage(caught: unknown): string {
  return caught instanceof Error ? caught.message : String(caught);
}
