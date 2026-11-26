import React, { useEffect, useRef, useState } from 'react';
import { Eraser, ImagePlus, PenLine, Check, X } from 'lucide-react';
import { compressCanvasToDataUrl, compressImageFileToDataUrl } from '../utils/portalHelpers';

interface SignaturePadModalProps {
  open: boolean;
  title: string;
  initialImage?: string;
  onClose: () => void;
  onSave: (dataUrl: string) => void;
}

export function SignaturePadModal({
  open,
  title,
  initialImage,
  onClose,
  onSave,
}: SignaturePadModalProps) {
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const [isDrawing, setIsDrawing] = useState(false);
  const [hasStroke, setHasStroke] = useState(false);
  const [uploadedPreview, setUploadedPreview] = useState<string>('');

  useEffect(() => {
    if (!open) return;
    setUploadedPreview(initialImage || '');
    setHasStroke(false);
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.lineWidth = 2.6;
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';
    ctx.strokeStyle = '#0f172a';
  }, [open, initialImage]);

  if (!open) return null;

  const getPoint = (e: React.PointerEvent<HTMLCanvasElement>) => {
    const canvas = canvasRef.current;
    if (!canvas) return { x: 0, y: 0 };
    const rect = canvas.getBoundingClientRect();
    const scaleX = canvas.width / rect.width;
    const scaleY = canvas.height / rect.height;
    return {
      x: (e.clientX - rect.left) * scaleX,
      y: (e.clientY - rect.top) * scaleY,
    };
  };

  const handlePointerDown = (e: React.PointerEvent<HTMLCanvasElement>) => {
    const canvas = canvasRef.current;
    const ctx = canvas?.getContext('2d');
    if (!canvas || !ctx) return;
    canvas.setPointerCapture(e.pointerId);
    const pt = getPoint(e);
    ctx.beginPath();
    ctx.moveTo(pt.x, pt.y);
    setIsDrawing(true);
    setHasStroke(true);
    setUploadedPreview('');
  };

  const handlePointerMove = (e: React.PointerEvent<HTMLCanvasElement>) => {
    if (!isDrawing) return;
    const canvas = canvasRef.current;
    const ctx = canvas?.getContext('2d');
    if (!canvas || !ctx) return;
    const pt = getPoint(e);
    ctx.lineTo(pt.x, pt.y);
    ctx.stroke();
  };

  const handlePointerUp = () => {
    setIsDrawing(false);
  };

  const handleClearCanvas = () => {
    const canvas = canvasRef.current;
    const ctx = canvas?.getContext('2d');
    if (canvas && ctx) {
      ctx.clearRect(0, 0, canvas.width, canvas.height);
    }
    setHasStroke(false);
    setUploadedPreview('');
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const compressed = await compressImageFileToDataUrl(file, 420, 200000);
    if (compressed) {
      setUploadedPreview(compressed);
      setHasStroke(false);
    }
  };

  const handleConfirmSave = () => {
    if (uploadedPreview) {
      onSave(uploadedPreview);
      onClose();
      return;
    }
    if (hasStroke && canvasRef.current) {
      onSave(compressCanvasToDataUrl(canvasRef.current, 190000));
      onClose();
      return;
    }
    onSave('');
    onClose();
  };

  return (
    <div className="no-print fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/60 backdrop-blur-xs">
      <div
        role="dialog"
        aria-modal="true"
        className="w-full max-w-lg rounded-xl bg-white border border-slate-200 shadow-xl overflow-hidden"
      >
        <div className="px-5 py-4 border-b border-slate-200 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <PenLine className="w-4 h-4 text-indigo-600" />
            <h3 className="text-sm sm:text-base font-bold text-slate-900">{title}</h3>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="p-1 rounded-lg text-slate-400 hover:text-slate-700"
            aria-label="Tutup"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        <div className="p-5 space-y-4">
          <p className="text-xs text-slate-500">
            Goreskan tanda tangan langsung pada kotak di bawah ini menggunakan mouse/sentuhan, atau unggah gambar PNG transparan.
          </p>

          <div className="relative rounded-xl border-2 border-dashed border-slate-300 bg-slate-50 overflow-hidden">
            {uploadedPreview ? (
              <div className="h-44 flex items-center justify-center p-4 bg-white">
                <img
                  src={uploadedPreview}
                  alt="Pratinjau Tanda Tangan"
                  className="max-h-full max-w-full object-contain"
                />
              </div>
            ) : (
              <canvas
                ref={canvasRef}
                width={460}
                height={180}
                onPointerDown={handlePointerDown}
                onPointerMove={handlePointerMove}
                onPointerUp={handlePointerUp}
                onPointerCancel={handlePointerUp}
                className="w-full h-44 bg-white touch-none cursor-crosshair block"
              />
            )}
          </div>

          <div className="flex flex-wrap items-center justify-between gap-2">
            <div className="flex items-center gap-2">
              <button
                type="button"
                onClick={handleClearCanvas}
                className="px-3 py-1.5 rounded-lg border border-slate-200 text-xs font-medium text-slate-700 hover:bg-slate-50 inline-flex items-center gap-1.5"
              >
                <Eraser className="w-3.5 h-3.5" />
                Bersihkan
              </button>
              <label className="cursor-pointer px-3 py-1.5 rounded-lg border border-indigo-200 bg-indigo-50 text-xs font-semibold text-indigo-700 hover:bg-indigo-100 inline-flex items-center gap-1.5">
                <ImagePlus className="w-3.5 h-3.5" />
                Unggah Gambar
                <input type="file" accept="image/*" className="hidden" onChange={handleFileUpload} />
              </label>
            </div>

            <div className="flex items-center gap-2">
              <button
                type="button"
                onClick={onClose}
                className="px-3.5 py-2 rounded-lg border border-slate-200 text-xs font-medium text-slate-700 hover:bg-slate-50"
              >
                Batal
              </button>
              <button
                type="button"
                onClick={handleConfirmSave}
                className="px-4 py-2 rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-semibold inline-flex items-center gap-1.5"
              >
                <Check className="w-3.5 h-3.5" />
                Simpan TTD
              </button>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
