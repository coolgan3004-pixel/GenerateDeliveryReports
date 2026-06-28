// Audio recorder — captures mic, down-samples to 16 kHz mono PCM, encodes WAV, returns base64.
// Used by the Meeting Minutes panel on the CSAT Report page.

window.audioRecorder = (() => {
    let audioContext = null;
    let mediaStream = null;
    let workletNode = null;
    let pcmChunks = [];
    let dotnetRef = null;

    const TARGET_SAMPLE_RATE = 16000;

    async function start(dotnetObjRef) {
        dotnetRef = dotnetObjRef;
        pcmChunks = [];

        mediaStream = await navigator.mediaDevices.getUserMedia({ audio: true, video: false });
        audioContext = new AudioContext({ sampleRate: TARGET_SAMPLE_RATE });

        // Register worklet only once per AudioContext
        await audioContext.audioWorklet.addModule('js/pcmProcessor.js');

        const source = audioContext.createMediaStreamSource(mediaStream);
        workletNode = new AudioWorkletNode(audioContext, 'pcm-processor');

        workletNode.port.onmessage = (e) => {
            pcmChunks.push(new Float32Array(e.data));
        };

        source.connect(workletNode);
        workletNode.connect(audioContext.destination);
    }

    async function stop() {
        if (workletNode) { workletNode.disconnect(); workletNode = null; }
        if (mediaStream) { mediaStream.getTracks().forEach(t => t.stop()); mediaStream = null; }
        if (audioContext) { await audioContext.close(); audioContext = null; }

        // Merge all PCM chunks
        const totalLen = pcmChunks.reduce((n, c) => n + c.length, 0);
        const merged = new Float32Array(totalLen);
        let offset = 0;
        for (const chunk of pcmChunks) { merged.set(chunk, offset); offset += chunk.length; }
        pcmChunks = [];

        // Encode as 16-bit PCM WAV
        const wavBytes = encodeWav(merged, TARGET_SAMPLE_RATE);
        const base64 = btoa(String.fromCharCode(...new Uint8Array(wavBytes)));
        return base64;
    }

    function encodeWav(samples, sampleRate) {
        const pcm16 = new Int16Array(samples.length);
        for (let i = 0; i < samples.length; i++)
            pcm16[i] = Math.max(-32768, Math.min(32767, Math.round(samples[i] * 32767)));

        const byteCount = pcm16.length * 2;
        const buf = new ArrayBuffer(44 + byteCount);
        const view = new DataView(buf);

        const write = (off, str) => { for (let i = 0; i < str.length; i++) view.setUint8(off + i, str.charCodeAt(i)); };

        write(0, 'RIFF');
        view.setUint32(4, 36 + byteCount, true);
        write(8, 'WAVE');
        write(12, 'fmt ');
        view.setUint32(16, 16, true);        // chunk size
        view.setUint16(20, 1, true);         // PCM
        view.setUint16(22, 1, true);         // mono
        view.setUint32(24, sampleRate, true);
        view.setUint32(28, sampleRate * 2, true);  // byte rate
        view.setUint16(32, 2, true);         // block align
        view.setUint16(34, 16, true);        // bits per sample
        write(36, 'data');
        view.setUint32(40, byteCount, true);

        const pcm8 = new Uint8Array(buf, 44);
        for (let i = 0; i < pcm16.length; i++) {
            pcm8[i * 2]     = pcm16[i] & 0xff;
            pcm8[i * 2 + 1] = (pcm16[i] >> 8) & 0xff;
        }
        return buf;
    }

    return { start, stop };
})();
