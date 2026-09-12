export default {
  async fetch(request) {
    const url = new URL(request.url);
    const APPS_SCRIPT_EXEC_URL = "https://script.google.com/macros/s/AKfycbxsWDpMsHQY5IFd1QSn7iw4iOMCQ-v19gGTiQvCqSJzOBhAoqkN8uq6CxP2WJ2kPU0d/exec"; // Ganti dengan URL Apps Script Anda

    // Default route: serve the Apps Script iframe
    const mainAppHtml = `
        <!DOCTYPE html>
        <html lang="id">
        <head>
          <meta charset="utf-8">
          <meta name="viewport" content="width=device-width, initial-scale=1">
          <link rel="icon" type="image/x-icon" href="https://i.ibb.co.com/ycqKLkNL/favicon-32x32.png">
          <script src="https://unpkg.com/html5-qrcode"></script>
          <title>Portal Kopinka</title>
          <style>
            body, html { margin: 0; padding: 0; height: 100%; overflow: hidden; background-color: #ffffff; }
            iframe { position: absolute; top: 0; left: 0; width: 100%; height: 100%; border: none; }
            
            /* Overlay Kamera Top-Level */
            #camera-overlay {
              display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%;
              background: rgba(0,0,0,0.9); z-index: 999999; justify-content: center; align-items: center;
              flex-direction: column; color: white; font-family: sans-serif;
            }
            .cam-container { width: 90%; max-width: 500px; background: #fff; padding: 20px; border-radius: 15px; text-align: center; color: #333; }
            video { width: 100%; border-radius: 10px; background: #000; margin-bottom: 15px; }
            .cam-btns { display: flex; gap: 10px; justify-content: center; }
            .btn { padding: 12px 20px; border: none; border-radius: 8px; cursor: pointer; font-weight: bold; color: white; }
            .btn-capture { background: #28a745; }
            .btn-close { background: #dc3545; }
            .btn-switch { background: #007bff; }
            #reader-worker { width: 100%; border-radius: 10px; overflow: hidden; margin-bottom: 15px; }
          </style>
        </head>
        <body>
          <iframe
            id="app-frame"
            src="${APPS_SCRIPT_EXEC_URL}"
            allow="camera *; microphone *; autoplay *; fullscreen *"
            allowfullscreen>
          </iframe>

          <div id="camera-overlay">
            <div class="cam-container">
              <h3 id="cam-title" style="margin-top:0">Kamera Verifikasi</h3>
              
              <div id="photo-area">
                <video id="webcam-video" autoplay playsinline></video>
                <canvas id="webcam-canvas" style="display:none"></canvas>
              </div>

              <div id="scanner-area" style="display:none;">
                <div id="reader-worker"></div>
              </div>

              <div class="cam-btns">
                <button id="btn-snap" class="btn btn-capture" onclick="takeSnapshot()">📸 AMBIL FOTO</button>
                <button id="btn-sw" class="btn btn-switch" onclick="switchCam()">🔄</button>
                <button class="btn btn-close" onclick="stopCam()">BATAL</button>
              </div>
            </div>
          </div>

          <script>
            const overlay = document.getElementById('camera-overlay');
            const video = document.getElementById('webcam-video');
            const canvas = document.getElementById('webcam-canvas');
            const iframe = document.getElementById('app-frame');
            let appsScriptWindow = null; // Simpan referensi window spesifik Apps Script
            let appsScriptOrigin = null; // Akan diisi setelah Apps Script mengirimkan origin-nya
            let cloudflareWorkerOrigin = window.location.origin; // Worker's own origin
            let stream = null;
            let html5QrCode = null;
            let facing = 'environment';

            // Listen pesan dari Apps Script
            window.addEventListener('message', (event) => {
              // 1. Abaikan pesan jika tidak memiliki format data yang valid (bukan pesan aplikasi kita)
              if (!event.data || typeof event.data !== 'object' || !event.data.type) return;
              
              // 2. Daftar tipe pesan yang valid untuk diproses di Worker. 
              // Ini akan mencegah Worker memproses pesan dari ekstensi browser (looping error).
              const validAppMessages = ['APPS_SCRIPT_READY', 'OPEN_CAMERA_REQUEST', 'OPEN_SCANNER_REQUEST', 'APPS_SCRIPT_ORIGIN'];
              if (!validAppMessages.includes(event.data.type)) return;

              // Prioritaskan menerima sinyal READY dari Apps Script untuk mendapatkan origin-nya
              if (event.data.type === 'APPS_SCRIPT_READY') {
                // Simpan origin Apps Script yang baru saja mengirim sinyal ready
                appsScriptOrigin = (event.origin === 'null' || !event.origin) ? 'null' : event.origin;
                appsScriptWindow = event.source; // Tangkap window yang mengirim sinyal READY
                const target = event.source || iframe.contentWindow;
                if (target) {
                  console.log('Cloudflare Worker: Mengirim WORKER_ORIGIN...');
                  target.postMessage({ type: 'WORKER_ORIGIN', origin: cloudflareWorkerOrigin }, '*');
                }
                return; // Jangan proses pesan lain, ini hanya untuk handshake awal
              }

              // Validasi Origin: Izinkan jika origin cocok, atau jika keduanya adalah 'null'
              const incomingOrigin = (event.origin === 'null' || !event.origin) ? 'null' : event.origin;
              if (appsScriptOrigin !== '*') {
                if (incomingOrigin !== 'null' && incomingOrigin !== appsScriptOrigin) {
                   console.warn('Cloudflare Worker: Dropping message dari origin tidak dikenal:', event.origin, event.data.type);
                   return;
                }
              }

              if (event.data.type === 'OPEN_CAMERA_REQUEST') {
                startCam();
              } else if (event.data.type === 'OPEN_SCANNER_REQUEST') {
                startScanner();
              } else if (event.data.type === 'APPS_SCRIPT_ORIGIN') {
                // Ini adalah konfirmasi dari Apps Script, origin sudah disimpan
                console.log('Cloudflare Worker: Final handshake complete dengan Apps Script origin:', appsScriptOrigin);
              }
            });

            async function startCam() {
              document.getElementById('cam-title').innerText = "Kamera Verifikasi";
              document.getElementById('photo-area').style.display = 'block';
              document.getElementById('scanner-area').style.display = 'none';
              document.getElementById('btn-snap').style.display = 'inline-block';
              document.getElementById('btn-sw').style.display = 'inline-block';
              
              overlay.style.display = 'flex';
              try {
                if(stream) stream.getTracks().forEach(t => t.stop());
                stream = await navigator.mediaDevices.getUserMedia({ 
                  video: { facingMode: facing, width: { ideal: 1280 }, height: { ideal: 720 } }, 
                  audio: false 
                });
                video.srcObject = stream;
                // Mirror jika kamera depan
                video.style.transform = facing === 'user' ? 'scaleX(-1)' : 'scaleX(1)';
              } catch (err) {
                alert("Kamera Error: " + err.message);
                stopCam();
              }
            }

            async function startScanner() {
              document.getElementById('cam-title').innerText = "Scan Barcode Voucher";
              document.getElementById('photo-area').style.display = 'none';
              document.getElementById('scanner-area').style.display = 'block';
              document.getElementById('btn-snap').style.display = 'none';
              document.getElementById('btn-sw').style.display = 'none';
              
              overlay.style.display = 'flex';
              
              if (!html5QrCode) {
                html5QrCode = new Html5Qrcode("reader-worker");
              }
              
              // Konfigurasi khusus untuk barcode batang (Code 39)
              const config = { 
                fps: 20, 
                qrbox: { width: 350, height: 150 }, // Lebih lebar untuk barcode linear
                aspectRatio: 1.0,
                formatsToSupport: [ Html5QrcodeSupportedFormats.CODE_39, Html5QrcodeSupportedFormats.CODE_128, Html5QrcodeSupportedFormats.QR_CODE ]
              };
              
              try {
                await html5QrCode.start({ facingMode: "environment" }, config, (decodedText) => {
                  // Kirim hasil scan ke Apps Script
                  const targetWindow = appsScriptWindow || iframe.contentWindow;
                  const targetOrigin = (appsScriptOrigin === 'null' || !appsScriptOrigin || appsScriptOrigin === '*') ? '*' : appsScriptOrigin;
                  targetWindow.postMessage({ type: 'SCANNER_RESULT', text: decodedText }, targetOrigin);
                  stopCam(true);
                });
              } catch (err) {
                alert("Scanner Error: " + err);
                stopCam();
              }
            }

            function stopCam(isCapture = false) {
              if(html5QrCode && html5QrCode.isScanning) {
                html5QrCode.stop().catch(err => console.error(err));
              }
              if(stream) stream.getTracks().forEach(t => t.stop());
              stream = null;
              overlay.style.display = 'none';
              // Kirim kabar ke Apps Script kalau batal
              if (!isCapture) {
                const targetWindow = appsScriptWindow || iframe.contentWindow;
                const target = (appsScriptOrigin === 'null' || !appsScriptOrigin) ? '*' : appsScriptOrigin;
                targetWindow.postMessage({ type: 'CAMERA_CANCELLED' }, target);
              }
            }

            function switchCam() {
              facing = facing === 'user' ? 'environment' : 'user';
              startCam();
            }

            function takeSnapshot() {
              if (!video.videoWidth) return alert("Video belum siap, silakan tunggu sebentar.");
              
              canvas.width = video.videoWidth;
              canvas.height = video.videoHeight;
              const ctx = canvas.getContext('2d');
              
              // Handle mirroring saat gambar
              if(facing === 'user') {
                ctx.translate(canvas.width, 0);
                ctx.scale(-1, 1);
              }
              
              ctx.drawImage(video, 0, 0);
              const base64 = canvas.toDataURL('image/jpeg', 0.8);

              // Gunakan '*' sebagai fallback jika origin tidak stabil di browser mobile
              const targetWindow = appsScriptWindow || iframe.contentWindow;
              const targetOrigin = (appsScriptOrigin === 'null' || !appsScriptOrigin || appsScriptOrigin === '*') ? '*' : appsScriptOrigin;
              // Kirim hasil ke Apps Script
              targetWindow.postMessage({ 
                type: 'CAMERA_RESULT', 
                base64: base64 
              }, targetOrigin);
              
              stopCam(true);
            }
          </script>
        </body>
        </html>
        `;

    return new Response(mainAppHtml, {
      headers: { 
        "content-type": "text/html;charset=UTF-8",
        "Permissions-Policy": "camera=(*), microphone=(*), autoplay=(*)",
        "Access-Control-Allow-Origin": "*"
      },
    });
  }
}
