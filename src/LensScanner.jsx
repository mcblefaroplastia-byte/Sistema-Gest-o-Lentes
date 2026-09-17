import React, { useEffect, useRef, useState } from 'react';
import { ScanLine, Camera, Keyboard, Search, X, CheckCircle2, AlertTriangle } from 'lucide-react';

/**
 * LensScanner.jsx
 * Componente independente para leitura de código de barras / QR.
 *
 * Funciona imediatamente com leitores USB/Bluetooth que se comportam como teclado.
 * Para câmera, usa BarcodeDetector quando disponível no navegador.
 *
 * Props:
 * - lenses: array de lentes. Procura por lens.barcode
 * - onSelect(lens, code): chamado quando encontra uma lente
 * - onUnknown(code): chamado quando o código ainda não está cadastrado
 * - title: título opcional
 */
export default function LensScanner({
  lenses = [],
  onSelect,
  onUnknown,
  title = 'Leitor de lentes',
}) {
  const [code, setCode] = useState('');
  const [status, setStatus] = useState(null);
  const [cameraOpen, setCameraOpen] = useState(false);
  const inputRef = useRef(null);
  const videoRef = useRef(null);
  const streamRef = useRef(null);
  const scanningRef = useRef(false);

  useEffect(() => {
    inputRef.current?.focus();
  }, []);

  useEffect(() => {
    return () => stopCamera();
  }, []);

  function normalize(value) {
    return String(value || '').trim();
  }

  function findLens(value) {
    const normalized = normalize(value);
    return lenses.find(
      (lens) => normalize(lens.barcode).toLowerCase() === normalized.toLowerCase()
    );
  }

  function processCode(value) {
    const normalized = normalize(value);
    if (!normalized) return;

    const lens = findLens(normalized);

    if (lens) {
      setStatus({
        type: 'success',
        text: `${lens.name || 'Lente'} ${lens.power || ''}`.trim(),
        lens,
      });
      onSelect?.(lens, normalized);
    } else {
      setStatus({
        type: 'warning',
        text: `Código ${normalized} ainda não está cadastrado.`,
      });
      onUnknown?.(normalized);
    }

    setCode('');
    setTimeout(() => inputRef.current?.focus(), 50);
  }

  function handleSubmit(e) {
    e.preventDefault();
    processCode(code);
  }

  async function startCamera() {
    setStatus(null);

    if (!('BarcodeDetector' in window)) {
      setStatus({
        type: 'warning',
        text: 'A leitura pela câmera não está disponível neste navegador. Use o leitor USB/Bluetooth ou digite o código.',
      });
      return;
    }

    try {
      const supported = await window.BarcodeDetector.getSupportedFormats();
      const wanted = [
        'qr_code',
        'code_128',
        'code_39',
        'ean_13',
        'ean_8',
        'upc_a',
        'upc_e',
        'itf',
        'data_matrix',
      ].filter((format) => supported.includes(format));

      const detector = new window.BarcodeDetector({
        formats: wanted.length ? wanted : supported,
      });

      const stream = await navigator.mediaDevices.getUserMedia({
        video: { facingMode: { ideal: 'environment' } },
        audio: false,
      });

      streamRef.current = stream;
      setCameraOpen(true);
      scanningRef.current = true;

      setTimeout(async () => {
        if (!videoRef.current) return;
        videoRef.current.srcObject = stream;
        await videoRef.current.play();

        const scan = async () => {
          if (!scanningRef.current || !videoRef.current) return;

          try {
            const results = await detector.detect(videoRef.current);
            if (results?.length) {
              const value = results[0].rawValue;
              stopCamera();
              processCode(value);
              return;
            }
          } catch (err) {
            console.error('Erro ao ler código:', err);
          }

          requestAnimationFrame(scan);
        };

        scan();
      }, 100);
    } catch (err) {
      console.error(err);
      stopCamera();
      setStatus({
        type: 'warning',
        text: 'Não foi possível abrir a câmera. Verifique a permissão do navegador ou use o leitor USB/Bluetooth.',
      });
    }
  }

  function stopCamera() {
    scanningRef.current = false;
    streamRef.current?.getTracks()?.forEach((track) => track.stop());
    streamRef.current = null;
    if (videoRef.current) videoRef.current.srcObject = null;
    setCameraOpen(false);
  }

  return (
    <div className="lens-scanner">
      <div className="lens-scanner-head">
        <div className="lens-scanner-icon">
          <ScanLine size={22} />
        </div>
        <div>
          <span>LEITURA RÁPIDA</span>
          <h3>{title}</h3>
          <p>Escaneie o código da embalagem para localizar a lente.</p>
        </div>
      </div>

      <form className="lens-scanner-form" onSubmit={handleSubmit}>
        <div className="lens-scanner-input">
          <Search size={18} />
          <input
            ref={inputRef}
            value={code}
            onChange={(e) => setCode(e.target.value)}
            placeholder="Escaneie ou digite o código..."
            autoComplete="off"
          />
        </div>

        <button type="submit" className="primary">
          <Keyboard size={16} />
          Localizar
        </button>

        <button
          type="button"
          className="secondary"
          onClick={cameraOpen ? stopCamera : startCamera}
        >
          {cameraOpen ? <X size={16} /> : <Camera size={16} />}
          {cameraOpen ? 'Fechar câmera' : 'Ler com câmera'}
        </button>
      </form>

      <div className="lens-scanner-hint">
        <Keyboard size={15} />
        <span>
          Leitor USB/Bluetooth: basta apontar para o código. O leitor envia o
          código como teclado e normalmente finaliza com Enter.
        </span>
      </div>

      {cameraOpen && (
        <div className="lens-camera">
          <video ref={videoRef} playsInline muted />
          <div className="lens-camera-frame">
            <ScanLine size={34} />
            <span>Posicione o código dentro da área</span>
          </div>
        </div>
      )}

      {status?.type === 'success' && (
        <div className="lens-scan-result success">
          <CheckCircle2 size={20} />
          <div>
            <b>Lente localizada</b>
            <span>{status.text}</span>
            {status.lens?.lot && <small>Lote: {status.lens.lot}</small>}
            {status.lens?.qty !== undefined && (
              <small>Saldo atual: {status.lens.qty}</small>
            )}
          </div>
        </div>
      )}

      {status?.type === 'warning' && (
        <div className="lens-scan-result warning">
          <AlertTriangle size={20} />
          <div>
            <b>Código não localizado</b>
            <span>{status.text}</span>
          </div>
        </div>
      )}
    </div>
  );
}
