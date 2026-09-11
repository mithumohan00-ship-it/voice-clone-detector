/**
 * DupDub Voice Cloning Studio — Interactive Core Logic
 * Implements:
 * 1. File Upload (drag-and-drop & file picker)
 * 2. Microphone Audio Recorder with Live Wave Visualizer
 * 3. Stepper Synthesis Pipeline Animation ("Wait a few seconds")
 * 4. Realistic Web Audio & Speech Synthesis Playback
 * 5. Interactive Waveform Player with Scrubbing
 * 6. Custom Text-to-Speech Generation & Download
 */

document.addEventListener('DOMContentLoaded', () => {
  // DOM Elements
  const uploadCard = document.getElementById('uploadCard');
  const fileInput = document.getElementById('fileInput');
  const uploadFeedback = document.getElementById('uploadFeedback');
  const fileNameDisplay = document.getElementById('fileNameDisplay');
  const btnRemoveFile = document.getElementById('btnRemoveFile');

  const recordCard = document.getElementById('recordCard');
  const micBtn = document.getElementById('micBtn');
  const recordingWidget = document.getElementById('recordingWidget');
  const recTimer = document.getElementById('recTimer');
  const btnStopRec = document.getElementById('btnStopRec');

  const node1 = document.getElementById('node1');
  const node2 = document.getElementById('node2');
  const lineProgress = document.getElementById('lineProgress');
  const connectorLabel = document.getElementById('connectorLabel');

  const resultCard = document.getElementById('resultCard');
  const audioPlayerBox = document.getElementById('audioPlayerBox');
  const playBtn = document.getElementById('playBtn');
  const iconPlay = playBtn ? playBtn.querySelector('.icon-play') : null;
  const iconPause = playBtn ? playBtn.querySelector('.icon-pause') : null;
  const waveformBarsContainer = document.getElementById('waveformBars');
  const audioTime = document.getElementById('audioTime');
  const holoOrb = document.getElementById('holoOrb');

  const btnTestVoice = document.getElementById('btnTestVoice');
  const btnDownloadVoice = document.getElementById('btnDownloadVoice');

  // State
  let isRecording = false;
  let recInterval = null;
  let recSeconds = 0;
  let mediaRecorder = null;
  let audioChunks = [];
  let isPlaying = false;
  let playInterval = null;
  let playTime = 0;
  const totalDuration = 8; // 8 seconds demo sample
  let audioCtx = null;
  let activeAudio = null;

  // -------------------------------------------------------------
  // 1. GENERATE WAVEFORM BARS
  // -------------------------------------------------------------
  const TOTAL_BARS = 36;
  const waveHeights = [
    18, 30, 45, 25, 60, 85, 70, 95, 40, 65, 80, 50, 
    90, 100, 75, 55, 80, 95, 65, 45, 70, 85, 90, 60, 
    40, 75, 85, 60, 45, 70, 55, 35, 50, 40, 25, 15
  ];

  function renderWaveform() {
    if (!waveformBarsContainer) return;
    waveformBarsContainer.innerHTML = '';
    for (let i = 0; i < TOTAL_BARS; i++) {
      const bar = document.createElement('div');
      bar.className = 'w-bar';
      const heightPercent = waveHeights[i % waveHeights.length];
      bar.style.height = `${Math.max(15, heightPercent)}%`;
      bar.dataset.index = i;

      // Click to scrub
      bar.addEventListener('click', (e) => {
        e.stopPropagation();
        const scrubPercent = i / TOTAL_BARS;
        seekToPercent(scrubPercent);
      });

      waveformBarsContainer.appendChild(bar);
    }
  }

  function updateWaveformProgress(percent) {
    if (!waveformBarsContainer) return;
    const bars = waveformBarsContainer.querySelectorAll('.w-bar');
    const activeCount = Math.floor(percent * bars.length);
    bars.forEach((bar, idx) => {
      if (idx <= activeCount) {
        bar.classList.add('active');
      } else {
        bar.classList.remove('active');
      }
    });
  }

  function seekToPercent(percent) {
    playTime = percent * totalDuration;
    updateTimeDisplay(playTime);
    updateWaveformProgress(percent);
    if (isPlaying) {
      restartAudioAt(playTime);
    }
  }

  function updateTimeDisplay(seconds) {
    if (!audioTime) return;
    const s = Math.floor(seconds);
    const formatted = `00:0${s}`.slice(-5);
    audioTime.textContent = formatted;
  }

  renderWaveform();

  // -------------------------------------------------------------
  // 2. SYNTHESIS PIPELINE ANIMATION ("Wait a few seconds")
  // -------------------------------------------------------------
  function triggerSynthesis(sourceName = 'audio sample') {
    // Reset Line
    if (lineProgress) {
      lineProgress.style.transition = 'width 2.6s cubic-bezier(0.4, 0, 0.2, 1)';
      lineProgress.style.width = '100%';
    }

    if (node1) node1.classList.remove('active-pulse');
    if (node2) node2.classList.remove('active-pulse');

    // Stepper Label sequence
    if (connectorLabel) {
      connectorLabel.style.color = '#2563EB';
      connectorLabel.innerHTML = 'Cloning<br>voice...';
    }

    // Vigorously animate Orb bars
    if (holoOrb) {
      holoOrb.style.transform = 'scale(1.12)';
      holoOrb.style.boxShadow = '0 25px 50px rgba(56, 189, 248, 0.6), 0 0 25px rgba(236, 72, 153, 0.4)';
    }

    setTimeout(() => {
      if (connectorLabel) {
        connectorLabel.innerHTML = 'Analyzing<br>timbre...';
      }
    }, 900);

    setTimeout(() => {
      if (connectorLabel) {
        connectorLabel.innerHTML = 'Synthesizing<br>nuances...';
      }
    }, 1800);

    setTimeout(() => {
      // Step 2 reached
      if (node2) {
        node2.classList.add('active-pulse');
        node2.style.background = '#0284C7';
        node2.style.color = '#FFFFFF';
      }
      if (connectorLabel) {
        connectorLabel.style.color = '#059669';
        connectorLabel.innerHTML = 'Voice<br>Cloned! ✨';
      }
      if (holoOrb) {
        holoOrb.style.transform = '';
        holoOrb.style.boxShadow = '';
      }

      // Smoothly expand / activate Result Audio Player
      if (audioPlayerBox) {
        audioPlayerBox.classList.add('show');
        audioPlayerBox.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
      }

      // Auto-play the preview voice!
      startAudioPlayback();
    }, 2800);
  }

  // -------------------------------------------------------------
  // 3. FILE UPLOAD HANDLING
  // -------------------------------------------------------------
  if (uploadCard) {
    uploadCard.addEventListener('click', (e) => {
      if (e.target === btnRemoveFile) return;
      fileInput.click();
    });

    uploadCard.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault();
        fileInput.click();
      }
    });

    // Drag and Drop
    ['dragenter', 'dragover'].forEach(eventName => {
      uploadCard.addEventListener(eventName, (e) => {
        e.preventDefault();
        e.stopPropagation();
        uploadCard.classList.add('drag-active');
      }, false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
      uploadCard.addEventListener(eventName, (e) => {
        e.preventDefault();
        e.stopPropagation();
        uploadCard.classList.remove('drag-active');
      }, false);
    });

    uploadCard.addEventListener('drop', (e) => {
      const dt = e.dataTransfer;
      const files = dt.files;
      if (files.length > 0) {
        handleSelectedFile(files[0]);
      }
    });
  }

  if (fileInput) {
    fileInput.addEventListener('change', (e) => {
      if (e.target.files && e.target.files.length > 0) {
        handleSelectedFile(e.target.files[0]);
      }
    });
  }

  function handleSelectedFile(file) {
    const validExtensions = ['mp3', 'wav', 'mp4', 'mov', 'm4a', 'aac', 'ogg'];
    const fileExt = file.name.split('.').pop().toLowerCase();

    if (!validExtensions.includes(fileExt) && !file.type.startsWith('audio/') && !file.type.startsWith('video/')) {
      alert('Please upload an audio or video file (mp3, wav, mp4, mov, m4a)');
      return;
    }

    if (fileNameDisplay) {
      const sizeMB = (file.size / (1024 * 1024)).toFixed(1);
      fileNameDisplay.textContent = `${file.name} (${sizeMB} MB)`;
    }
    if (uploadFeedback) {
      uploadFeedback.classList.add('active');
    }

    // Trigger cloning workflow
    triggerSynthesis(file.name);
  }

  if (btnRemoveFile) {
    btnRemoveFile.addEventListener('click', (e) => {
      e.stopPropagation();
      if (uploadFeedback) uploadFeedback.classList.remove('active');
      if (fileInput) fileInput.value = '';
      resetStepper();
    });
  }

  // -------------------------------------------------------------
  // 4. MICROPHONE VOICE RECORDING
  // -------------------------------------------------------------
  if (recordCard) {
    recordCard.addEventListener('click', (e) => {
      if (e.target === btnStopRec || e.target.closest('#recordingWidget')) return;
      toggleRecording();
    });

    recordCard.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault();
        toggleRecording();
      }
    });
  }

  if (btnStopRec) {
    btnStopRec.addEventListener('click', (e) => {
      e.stopPropagation();
      stopRecording();
    });
  }

  function toggleRecording() {
    if (!isRecording) {
      startRecording();
    } else {
      stopRecording();
    }
  }

  async function startRecording() {
    isRecording = true;
    recSeconds = 0;
    if (recTimer) recTimer.textContent = '00:00';
    if (recordingWidget) recordingWidget.classList.add('active');
    if (recordCard) recordCard.classList.add('recording-active');

    // Try Real Web Audio MediaStream
    try {
      if (navigator.mediaDevices && navigator.mediaDevices.getUserMedia) {
        const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
        mediaRecorder = new MediaRecorder(stream);
        audioChunks = [];

        mediaRecorder.ondataavailable = (event) => {
          if (event.data.size > 0) audioChunks.push(event.data);
        };

        mediaRecorder.onstop = () => {
          stream.getTracks().forEach(track => track.stop());
        };

        mediaRecorder.start();
      }
    } catch (err) {
      console.log('Microphone access simulated or not granted:', err);
    }

    // Timer countdown
    recInterval = setInterval(() => {
      recSeconds++;
      if (recTimer) {
        recTimer.textContent = `00:0${recSeconds}`.slice(-5);
      }
      if (recSeconds >= 5) {
        stopRecording();
      }
    }, 1000);
  }

  function stopRecording() {
    if (!isRecording) return;
    isRecording = false;
    clearInterval(recInterval);

    if (recordingWidget) recordingWidget.classList.remove('active');
    if (recordCard) recordCard.classList.remove('recording-active');

    if (mediaRecorder && mediaRecorder.state !== 'inactive') {
      try { mediaRecorder.stop(); } catch(e) {}
    }

    // Trigger synthesis
    triggerSynthesis('recorded voice sample');
  }

  function resetStepper() {
    if (lineProgress) {
      lineProgress.style.transition = 'width 0.3s ease';
      lineProgress.style.width = '0%';
    }
    if (node1) {
      node1.className = 'step-node node-1 active';
    }
    if (node2) {
      node2.className = 'step-node node-2';
      node2.style.background = '';
      node2.style.color = '';
    }
    if (connectorLabel) {
      connectorLabel.style.color = '';
      connectorLabel.innerHTML = 'Wait a few<br>seconds';
    }
  }

  // -------------------------------------------------------------
  // 5. AUDIO PLAYBACK & SYNTHESIS ENGINE
  // -------------------------------------------------------------
  function getAudioContext() {
    if (!audioCtx) {
      const AudioCtxClass = window.AudioContext || window.webkitAudioContext;
      if (AudioCtxClass) {
        audioCtx = new AudioCtxClass();
      }
    }
    if (audioCtx && audioCtx.state === 'suspended') {
      audioCtx.resume();
    }
    return audioCtx;
  }

  // Plays a rich, natural melodic speech tone using Web Audio synthesis
  function playSyntheticVoice(startOffset = 0) {
    const ctx = getAudioContext();
    if (!ctx) return null;

    const osc = ctx.createOscillator();
    const formantFilter1 = ctx.createBiquadFilter();
    const formantFilter2 = ctx.createBiquadFilter();
    const gainNode = ctx.createGain();

    osc.type = 'sawtooth';
    // Human voice fundamental frequency pitch melody
    const now = ctx.currentTime;
    const baseFreq = 145; // Natural Baritone/Tenor speech pitch
    osc.frequency.setValueAtTime(baseFreq, now);
    osc.frequency.exponentialRampToValueAtTime(175, now + 1.2);
    osc.frequency.exponentialRampToValueAtTime(130, now + 2.5);
    osc.frequency.exponentialRampToValueAtTime(160, now + 4.0);
    osc.frequency.exponentialRampToValueAtTime(125, now + 5.5);
    osc.frequency.exponentialRampToValueAtTime(150, now + 7.0);

    // Vocal Formant F1 (around 600Hz, vowel "ah/oh")
    formantFilter1.type = 'bandpass';
    formantFilter1.frequency.setValueAtTime(650, now);
    formantFilter1.Q.value = 4.5;

    // Vocal Formant F2 (around 1200Hz)
    formantFilter2.type = 'bandpass';
    formantFilter2.frequency.setValueAtTime(1350, now);
    formantFilter2.Q.value = 4.0;

    gainNode.gain.setValueAtTime(0.001, now);
    gainNode.gain.linearRampToValueAtTime(0.22, now + 0.15);
    gainNode.gain.setValueAtTime(0.22, now + 7.5);
    gainNode.gain.linearRampToValueAtTime(0.001, now + totalDuration);

    osc.connect(formantFilter1);
    formantFilter1.connect(formantFilter2);
    formantFilter2.connect(gainNode);
    gainNode.connect(ctx.destination);

    osc.start(now);
    osc.stop(now + (totalDuration - startOffset));

    return {
      stop: () => {
        try {
          gainNode.gain.linearRampToValueAtTime(0.001, ctx.currentTime + 0.05);
          setTimeout(() => osc.stop(), 60);
        } catch(e) {}
      }
    };
  }

  // Browser SpeechSynthesis for natural voice reading
  function speakVoiceSample(text = "Hello! This voice was cloned authentically in just seconds with DupDub's neural engine.") {
    if ('speechSynthesis' in window) {
      window.speechSynthesis.cancel();
      const utterance = new SpeechSynthesisUtterance(text);
      utterance.rate = 1.02;
      utterance.pitch = 1.0;
      
      // Select best natural voice
      const voices = window.speechSynthesis.getVoices();
      if (voices.length > 0) {
        const preferredVoice = voices.find(v => v.lang.includes('en') && (v.name.includes('Natural') || v.name.includes('Google') || v.name.includes('Premium')));
        if (preferredVoice) utterance.voice = preferredVoice;
      }
      
      utterance.onend = () => {
        stopAudioPlayback();
      };
      
      window.speechSynthesis.speak(utterance);
    }
  }

  function startAudioPlayback() {
    isPlaying = true;
    if (iconPlay) iconPlay.style.display = 'none';
    if (iconPause) iconPause.style.display = 'block';

    // Start synthesized harmonic sound
    activeAudio = playSyntheticVoice(playTime);
    speakVoiceSample();

    clearInterval(playInterval);
    playInterval = setInterval(() => {
      playTime += 0.2;
      if (playTime >= totalDuration) {
        stopAudioPlayback();
        playTime = 0;
        updateTimeDisplay(0);
        updateWaveformProgress(0);
        return;
      }
      updateTimeDisplay(playTime);
      updateWaveformProgress(playTime / totalDuration);
    }, 200);
  }

  function stopAudioPlayback() {
    isPlaying = false;
    clearInterval(playInterval);
    if (iconPlay) iconPlay.style.display = 'block';
    if (iconPause) iconPause.style.display = 'none';

    if (activeAudio) {
      activeAudio.stop();
      activeAudio = null;
    }
    if ('speechSynthesis' in window) {
      window.speechSynthesis.cancel();
    }
  }

  function restartAudioAt(seconds) {
    stopAudioPlayback();
    playTime = seconds;
    startAudioPlayback();
  }

  if (playBtn) {
    playBtn.addEventListener('click', () => {
      if (!isPlaying) {
        startAudioPlayback();
      } else {
        stopAudioPlayback();
      }
    });
  }

  // Click on Result Card triggers preview if player is hidden
  if (resultCard) {
    resultCard.addEventListener('click', (e) => {
      if (e.target.closest('.audio-player-box') || e.target.closest('button')) return;
      if (audioPlayerBox && !audioPlayerBox.classList.contains('show')) {
        audioPlayerBox.classList.add('show');
        startAudioPlayback();
      }
    });
  }

  // -------------------------------------------------------------
  // 6. CUSTOM TEXT-TO-SPEECH GENERATOR
  // -------------------------------------------------------------
  if (btnTestVoice) {
    btnTestVoice.addEventListener('click', () => {
      const customText = prompt(
        'Enter text to synthesize with your cloned voice:',
        'DupDub captures every vocal timbre, emotional accent, and subtle inflection perfectly!'
      );
      if (customText && customText.trim()) {
        stopAudioPlayback();
        speakVoiceSample(customText.trim());
        // Animate waveform
        let t = 0;
        const dur = Math.max(4, Math.min(12, customText.length * 0.08));
        const intv = setInterval(() => {
          t += 0.2;
          updateWaveformProgress(t / dur);
          updateTimeDisplay(t);
          if (t >= dur) {
            clearInterval(intv);
            updateWaveformProgress(1);
          }
        }, 200);
      }
    });
  }

  // -------------------------------------------------------------
  // 7. DOWNLOAD AUDIO SAMPLE (WAV GENERATOR)
  // -------------------------------------------------------------
  if (btnDownloadVoice) {
    btnDownloadVoice.addEventListener('click', () => {
      generateAndDownloadWav();
    });
  }

  function generateAndDownloadWav() {
    const sampleRate = 44100;
    const duration = 5;
    const numFrames = sampleRate * duration;
    const buffer = new ArrayBuffer(44 + numFrames * 2);
    const view = new DataView(buffer);

    // Write WAV Header
    writeString(view, 0, 'RIFF');
    view.setUint32(4, 36 + numFrames * 2, true);
    writeString(view, 8, 'WAVE');
    writeString(view, 12, 'fmt ');
    view.setUint32(16, 16, true);
    view.setUint16(20, 1, true); // PCM format
    view.setUint16(22, 1, true); // Mono
    view.setUint32(24, sampleRate, true);
    view.setUint32(28, sampleRate * 2, true);
    view.setUint16(32, 2, true);
    view.setUint16(34, 16, true);
    writeString(view, 36, 'data');
    view.setUint32(40, numFrames * 2, true);

    // Fill harmonic voice PCM data
    let offset = 44;
    for (let i = 0; i < numFrames; i++) {
      const t = i / sampleRate;
      const freq = 150 + Math.sin(t * 3.5) * 25;
      const sample = Math.sin(2 * Math.PI * freq * t) * 0.4 
                   + Math.sin(2 * Math.PI * (freq * 2) * t) * 0.2
                   + Math.sin(2 * Math.PI * (freq * 3) * t) * 0.1;
      const val = Math.max(-1, Math.min(1, sample)) * 0x7FFF;
      view.setInt16(offset, val, true);
      offset += 2;
    }

    const blob = new Blob([view], { type: 'audio/wav' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'dupdub_cloned_voice.wav';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  function writeString(view, offset, string) {
    for (let i = 0; i < string.length; i++) {
      view.setUint8(offset + i, string.charCodeAt(i));
    }
  }

  // -------------------------------------------------------------
  // 8. INTERACTIVE HOLOGRAPHIC ORB
  // -------------------------------------------------------------
  if (holoOrb) {
    holoOrb.addEventListener('click', () => {
      // Toggle play preview
      if (!isPlaying) {
        startAudioPlayback();
      } else {
        stopAudioPlayback();
      }
    });
  }
});
