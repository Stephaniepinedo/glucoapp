22:17:24.527 Running build in Washington, D.C., USA (East) – iad1
22:17:24.527 Build machine configuration: 2 cores, 8 GB
22:17:24.629 Cloning github.com/Stephaniepinedo/glucoapp (Branch: main, Commit: 5b84314)
22:17:24.850 Cloning completed: 220.000ms
22:17:25.055 Restored build cache from previous deployment (G1q98eDFX1VsKCJovfbcvcjp2vN1)
22:17:25.253 Running "vercel build"
22:17:25.272 Vercel CLI 55.0.0
22:17:25.993 Installing dependencies...
22:17:28.815 
22:17:28.815 up to date in 3s
22:17:28.815 
22:17:28.816 7 packages are looking for funding
22:17:28.816   run `npm fund` for details
22:17:28.852 Running "npm run build"
22:17:29.296 
22:17:29.297 > glucoapp@1.0.0 build
22:17:29.297 > vite build
22:17:29.297 
22:17:29.495 The CJS build of Vite's Node API is deprecated. See https://vite.dev/guide/troubleshooting.html#vite-cjs-node-api-deprecated for more details.
22:17:29.548 vite v5.4.21 building for production...
22:17:29.604 transforming...
22:17:29.703 ✓ 6 modules transformed.
22:17:29.704 x Build failed in 128ms
22:17:29.705 error during build:
22:17:29.705 [vite:esbuild] Transform failed with 3 errors:
22:17:29.705 /vercel/path0/src/App.jsx:258:6: ERROR: The symbol "getToken" has already been declared
22:17:29.705 /vercel/path0/src/App.jsx:259:6: ERROR: The symbol "setToken" has already been declared
22:17:29.705 /vercel/path0/src/App.jsx:260:6: ERROR: The symbol "clearToken" has already been declared
22:17:29.705 file: /vercel/path0/src/App.jsx:258:6
22:17:29.706 
22:17:29.706 The symbol "getToken" has already been declared
22:17:29.706 256|  };
22:17:29.706 257|  
22:17:29.706 258|  const getToken = () => sessionStorage.getItem("ms_token");
22:17:29.706    |        ^
22:17:29.707 259|  const setToken = (t) => sessionStorage.setItem("ms_token", t);
22:17:29.707 260|  const clearToken = () => sessionStorage.removeItem("ms_token");
22:17:29.707 
22:17:29.707 The symbol "setToken" has already been declared
22:17:29.707 257|  
22:17:29.707 258|  const getToken = () => sessionStorage.getItem("ms_token");
22:17:29.707 259|  const setToken = (t) => sessionStorage.setItem("ms_token", t);
22:17:29.708    |        ^
22:17:29.708 260|  const clearToken = () => sessionStorage.removeItem("ms_token");
22:17:29.708 261|  
22:17:29.708 
22:17:29.708 The symbol "clearToken" has already been declared
22:17:29.708 258|  const getToken = () => sessionStorage.getItem("ms_token");
22:17:29.709 259|  const setToken = (t) => sessionStorage.setItem("ms_token", t);
22:17:29.709 260|  const clearToken = () => sessionStorage.removeItem("ms_token");
22:17:29.709    |        ^
22:17:29.709 261|  
22:17:29.709 262|  const msLogin = (onToken) => {
22:17:29.709 
22:17:29.710     at failureErrorWithLog (/vercel/path0/node_modules/esbuild/lib/main.js:1472:15)
22:17:29.710     at /vercel/path0/node_modules/esbuild/lib/main.js:755:50
22:17:29.710     at responseCallbacks.<computed> (/vercel/path0/node_modules/esbuild/lib/main.js:622:9)
22:17:29.711     at handleIncomingPacket (/vercel/path0/node_modules/esbuild/lib/main.js:677:12)
22:17:29.711     at Socket.readFromStdout (/vercel/path0/node_modules/esbuild/lib/main.js:600:7)
22:17:29.711     at Socket.emit (node:events:509:28)
22:17:29.711     at addChunk (node:internal/streams/readable:563:12)
22:17:29.712     at readableAddChunkPushByteMode (node:internal/streams/readable:514:3)
22:17:29.712     at Readable.push (node:internal/streams/readable:394:5)
22:17:29.712     at Pipe.onStreamRead (node:internal/stream_base_commons:189:23)
22:17:29.733 Error: Command "npm run build" exited with 1
