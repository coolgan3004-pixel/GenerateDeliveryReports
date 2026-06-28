// AudioWorklet processor — collects raw PCM float32 samples and posts them back.
class PcmProcessor extends AudioWorkletProcessor {
    process(inputs) {
        const input = inputs[0];
        if (input && input[0]) {
            this.port.postMessage(input[0].buffer, [input[0].buffer]);
        }
        return true;
    }
}

registerProcessor('pcm-processor', PcmProcessor);
