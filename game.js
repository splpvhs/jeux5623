var gk_isXlsx = false;
        var gk_xlsxFileLookup = {};
        var gk_fileData = {};
        function filledCell(cell) {
          return cell !== '' && cell != null;
        }
        function loadFileData(filename) {
        if (gk_isXlsx && gk_xlsxFileLookup[filename]) {
            try {
                var workbook = XLSX.read(gk_fileData[filename], { type: 'base64' });
                var firstSheetName = workbook.SheetNames[0];
                var worksheet = workbook.Sheets[firstSheetName];

                // Convert sheet to JSON to filter blank rows
                var jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, blankrows: false, defval: '' });
                // Filter out blank rows (rows where all cells are empty, null, or undefined)
                var filteredData = jsonData.filter(row => row.some(filledCell));

                // Heuristic to find the header row by ignoring rows with fewer filled cells than the next row
                var headerRowIndex = filteredData.findIndex((row, index) =>
                  row.filter(filledCell).length >= filteredData[index + 1]?.filter(filledCell).length
                );
                // Fallback
                if (headerRowIndex === -1 || headerRowIndex > 25) {
                  headerRowIndex = 0;
                }

                // Convert filtered JSON back to CSV
                var csv = XLSX.utils.aoa_to_sheet(filteredData.slice(headerRowIndex)); // Create a new sheet from filtered array of arrays
                csv = XLSX.utils.sheet_to_csv(csv, { header: 1 });
                return csv;
            } catch (e) {
                console.error(e);
                return "";
            }
        }
        return gk_fileData[filename] || "";
        }

// تهيئة المتغيرات العامة
        let currentNumber = null;
        let currentMode = null;
        let questions = [];
        let currentQuestion = 0;
        let score = 0;
        let mistakes = [];
        let answeredOperations = [];
        let soundEnabled = true;
        let musicEnabled = false;
        let currentMusicTrack = 1;
        let audioContext = null;
        let generatedMusicInterval = null;
        let generatedMusicNodes = [];
        let musicFallback = false;
        const tableColors = ['dark', 'purple', 'navy', 'bordeaux', 'olive'];
        let allErrors = [];
        let celebrationDuckCount = 0;
        let celebrationPrevVolumes = [];
        try {
            allErrors = JSON.parse(localStorage.getItem('multiplicationErrors')) || [];
        } catch (e) {
            allErrors = [];
        }
        function saveErrors() {
            try {
                localStorage.setItem('multiplicationErrors', JSON.stringify(allErrors));
            } catch (e) {
                console.warn('LocalStorage inaccessible', e);
            }
        }
        window.addEventListener('beforeunload', saveErrors);

        // إنشاء أزرار الأرقام من 1 إلى 10
        function initNumberGrid() {
            const grid = document.getElementById('numberGrid');
            for (let i = 1; i <= 10; i++) {
                const btn = document.createElement('button');
                btn.className = 'number-btn';
                btn.textContent = i;
                btn.onclick = () => selectNumber(i);
                grid.appendChild(btn);
            }
        }

        // اختيار رقم من الشبكة
        function selectNumber(number) {
            currentNumber = number;
            document.getElementById('mainMenu').style.display = 'none';
            document.getElementById('optionsMenu').style.display = 'block';
            document.getElementById('selectedNumber').textContent = `جدول الضرب للرقم ${number}`;
        }

        function getAudioContext() {
            if (!audioContext) {
                audioContext = new (window.AudioContext || window.webkitAudioContext)();
            }
            return audioContext;
        }

        function playTone(frequency, type, duration) {
            if (!soundEnabled) return;
            const ctx = getAudioContext();
            const oscillator = ctx.createOscillator();
            const gain = ctx.createGain();
            oscillator.type = type;
            oscillator.frequency.value = frequency;
            oscillator.connect(gain);
            gain.connect(ctx.destination);
            gain.gain.setValueAtTime(0.0001, ctx.currentTime);
            gain.gain.exponentialRampToValueAtTime(0.25, ctx.currentTime + 0.01);
            oscillator.start();
            gain.gain.exponentialRampToValueAtTime(0.0001, ctx.currentTime + duration);
            oscillator.stop(ctx.currentTime + duration + 0.02);
        }

        function playToneSequence(sequence) {
            if (!soundEnabled) return;
            let start = getAudioContext().currentTime;
            sequence.forEach(note => {
                const ctx = getAudioContext();
                const oscillator = ctx.createOscillator();
                const gain = ctx.createGain();
                oscillator.type = note.type;
                oscillator.frequency.value = note.frequency;
                oscillator.connect(gain);
                gain.connect(ctx.destination);
                gain.gain.setValueAtTime(0.0001, start);
                gain.gain.exponentialRampToValueAtTime(0.25, start + 0.01);
                oscillator.start(start);
                gain.gain.exponentialRampToValueAtTime(0.0001, start + note.duration);
                oscillator.stop(start + note.duration + 0.02);
                start += note.duration + 0.04;
            });
        }

        function playCorrectSound() {
            if (!soundEnabled) return;
            const audioEl = document.getElementById('vraiAudio');
            if (audioEl) {
                try {
                    audioEl.currentTime = 0;
                    const p = audioEl.play();
                    if (p !== undefined) {
                        p.catch(err => {
                            console.warn('vraiAudio play failed, fallback to tone:', err);
                            playToneSequence([
                                { frequency: 880, type: 'sine', duration: 0.12 },
                                { frequency: 1040, type: 'triangle', duration: 0.14 }
                            ]);
                        });
                    }
                    return;
                } catch (e) {
                    console.warn('vraiAudio exception, fallback to tone:', e);
                }
            }
            playToneSequence([
                { frequency: 880, type: 'sine', duration: 0.12 },
                { frequency: 1040, type: 'triangle', duration: 0.14 }
            ]);
        }

        function playWrongSound() {
            if (!soundEnabled) return;
            const audioEl = document.getElementById('fauxAudio');
            if (audioEl) {
                try {
                    audioEl.currentTime = 0;
                    const p = audioEl.play();
                    if (p !== undefined) {
                        p.catch(err => {
                            console.warn('fauxAudio play failed, fallback to tone:', err);
                            playToneSequence([
                                { frequency: 330, type: 'sawtooth', duration: 0.16 },
                                { frequency: 220, type: 'square', duration: 0.18 },
                                { frequency: 180, type: 'triangle', duration: 0.14 }
                            ]);
                        });
                    }
                    return;
                } catch (e) {
                    console.warn('fauxAudio exception, fallback to tone:', e);
                }
            }
            playToneSequence([
                { frequency: 330, type: 'sawtooth', duration: 0.16 },
                { frequency: 220, type: 'square', duration: 0.18 },
                { frequency: 180, type: 'triangle', duration: 0.14 }
            ]);
        }

        function playCelebrationSound() {
            if (!soundEnabled) return;
            // Duck background audio volumes (do not pause)
            const duckFactor = 0.25; // volume multiplier during celebration
            const bgElements = [];
            // collect bgMusicTracks (Audio objects created in initBackgroundMusic)
            bgMusicTracks.forEach(t => { if (t) bgElements.push(t); });
            // collect any <audio id="bgMusicX"> elements in the DOM
            document.querySelectorAll('audio[id^="bgMusic"]').forEach(a => { if (a) bgElements.push(a); });
            // dedupe
            const uniqueBg = Array.from(new Set(bgElements));
            // if first ducking, save previous volumes
            if (celebrationDuckCount === 0) {
                celebrationPrevVolumes = uniqueBg.map(el => ({ el, vol: typeof el.volume === 'number' ? el.volume : 1 }));
                uniqueBg.forEach(({ volume }, i) => {
                    try { uniqueBg[i].volume = (typeof volume === 'number' ? volume : 1) * duckFactor; } catch (_) {}
                });
            }
            celebrationDuckCount++;

            const restoreVolumes = () => {
                celebrationDuckCount = Math.max(0, celebrationDuckCount - 1);
                if (celebrationDuckCount > 0) return; // wait for nested celebrations to finish
                celebrationPrevVolumes.forEach(item => {
                    try { item.el.volume = item.vol; } catch (_) {}
                });
                celebrationPrevVolumes = [];
            };

            const audioEl = document.getElementById('celebrationAudio');
            const fallbackSequence = [
                { frequency: 660, type: 'sine', duration: 0.12 },
                { frequency: 880, type: 'triangle', duration: 0.12 },
                { frequency: 1040, type: 'square', duration: 0.16 }
            ];

            if (audioEl) {
                try {
                    audioEl.loop = false;
                    audioEl.currentTime = 0;
                    const playPromise = audioEl.play();
                    if (playPromise !== undefined) {
                        playPromise.then(() => {
                            // played successfully, restore when ended
                            audioEl.addEventListener('ended', function onEnd() {
                                audioEl.removeEventListener('ended', onEnd);
                                restoreVolumes();
                            });
                        }).catch(err => {
                            console.warn('Celebration audio failed to play, falling back to tones:', err);
                            // fallback to tones and restore after estimated duration
                            const total = fallbackSequence.reduce((s, n) => s + n.duration + 0.04, 0);
                            playToneSequence(fallbackSequence);
                            setTimeout(restoreVolumes, (total + 0.05) * 1000);
                        });
                    } else {
                        // play returned undefined, still restore on ended
                        audioEl.addEventListener('ended', function onEnd() {
                            audioEl.removeEventListener('ended', onEnd);
                            restoreVolumes();
                        });
                    }
                    return;
                } catch (e) {
                    console.warn('Error playing celebration audio, fallback:', e);
                }
            }

            // fallback: play tone sequence and restore after its duration
            const total = fallbackSequence.reduce((s, n) => s + n.duration + 0.04, 0);
            playToneSequence(fallbackSequence);
            setTimeout(restoreVolumes, (total + 0.05) * 1000);
        }

        function toggleTheme() {
            const body = document.body;
            body.classList.toggle('dark-mode');
            document.getElementById('themeToggle').textContent = body.classList.contains('dark-mode') ? '☀️' : '🌙';
        }

        function toggleSound() {
            soundEnabled = !soundEnabled;
            document.getElementById('soundToggle').textContent = soundEnabled ? '🔊' : '🔇';
        }

        function getCurrentMusicElement() {
            return bgMusicTracks[currentMusicTrack] || document.getElementById(`bgMusic${currentMusicTrack}`);
        }

        function startGeneratedMusic(track) {
            stopGeneratedMusic();
            const ctx = getAudioContext();
            const patterns = {
                1: [[262, 330, 392], [330, 392, 523]],
                2: [[196, 247, 311], [247, 311, 392]],
                3: [[220, 277, 330], [277, 330, 392]],
                4: [[240, 320, 400], [320, 400, 480]],
                5: [[260, 310, 370], [310, 370, 440]]
            }[track] || [[220, 277, 330], [277, 330, 392]];
            let step = 0;
            generatedMusicInterval = setInterval(() => {
                const notes = patterns[step % patterns.length];
                notes.forEach((freq, index) => {
                    const osc = ctx.createOscillator();
                    const gain = ctx.createGain();
                    osc.type = index === 0 ? 'triangle' : index === 1 ? 'sine' : 'square';
                    osc.frequency.value = freq;
                    osc.connect(gain);
                    gain.connect(ctx.destination);
                    gain.gain.setValueAtTime(0.0001, ctx.currentTime);
                    gain.gain.exponentialRampToValueAtTime(0.18, ctx.currentTime + 0.02);
                    const duration = 0.4;
                    osc.start(ctx.currentTime);
                    gain.gain.exponentialRampToValueAtTime(0.0001, ctx.currentTime + duration);
                    osc.stop(ctx.currentTime + duration + 0.02);
                    generatedMusicNodes.push(osc);
                });
                step += 1;
            }, 700);
        }
        function stopGeneratedMusic() {
            if (generatedMusicInterval) {
                clearInterval(generatedMusicInterval);
                generatedMusicInterval = null;
            }
            generatedMusicNodes.forEach(osc => {
                try { osc.stop(); } catch (_) {}
            });
            generatedMusicNodes = [];
        }
        const bgMusicTracks = [];
        const bgMusicSources = ['music1.mp3', 'music2.mp3', 'music3.mp3', 'music4.mp3', 'music5.mp3'];

        function initBackgroundMusic() {
            bgMusicSources.forEach((src, index) => {
                const audio = new Audio(src);
                audio.loop = true;
                audio.preload = 'auto';
                audio.playsInline = true;
                audio.volume = 0.75;
                bgMusicTracks[index + 1] = audio;
            });
        }

        function stopAllBackgroundMusic() {
            bgMusicTracks.forEach(track => {
                if (!track) return;
                try {
                    track.pause();
                    track.currentTime = 0;
                } catch (_){ }
            });
            const audioTracks = document.querySelectorAll('audio[id^="bgMusic"]');
            audioTracks.forEach(track => {
                try {
                    track.pause();
                    track.currentTime = 0;
                } catch (_){ }
            });
        }

        function startBackgroundMusic() {
            if (!musicEnabled) return;
            const music = getCurrentMusicElement();
            if (!music) return;
            stopAllBackgroundMusic();
            stopGeneratedMusic();
            music.muted = false;
            music.volume = 0.75;
            music.currentTime = 0;
            const playPromise = music.play();
            if (playPromise !== undefined) {
                playPromise.catch(error => {
                    console.warn('Background music failed to play:', error);
                });
            }
        }

        function stopBackgroundMusic() {
            stopAllBackgroundMusic();
            stopGeneratedMusic();
        }

        function toggleMusic() {
            musicEnabled = !musicEnabled;
            document.getElementById('musicToggle').textContent = musicEnabled ? '▶️' : '⏹️';
            if (musicEnabled) {
                stopBackgroundMusic();
                startBackgroundMusic();
            } else {
                stopBackgroundMusic();
            }
        }

        function switchMusicTrack() {
            currentMusicTrack = currentMusicTrack === 5 ? 1 : currentMusicTrack + 1;
            document.getElementById('switchMusicBtn').textContent = `${currentMusicTrack}`;
            if (musicEnabled) {
                stopAllBackgroundMusic();
                const newMusic = getCurrentMusicElement();
                if (newMusic) {
                    newMusic.muted = false;
                    newMusic.volume = 0.75;
                    newMusic.currentTime = 0;
                    const playPromise = newMusic.play();
                    if (playPromise !== undefined) {
                        playPromise.catch(error => {
                            console.warn('Background music failed to switch:', error);
                        });
                    }
                }
            }
        }

        function showTable() {
            document.getElementById('optionsMenu').style.display = 'none';
            document.getElementById('tableView').style.display = 'block';
            document.getElementById('tableTitle').textContent = `جدول الضرب للرقم ${currentNumber}`;
            const content = document.getElementById('tableContent');
            content.innerHTML = '';
            for (let i = 0; i <= 10; i++) {
                const colorClass = tableColors[i % tableColors.length];
                const value = currentNumber * i;
                content.innerHTML += `<div class="table-row ${colorClass}">${currentNumber} × ${i} = ${value}</div>`;
            }
        }

        // بدء الاختبار (سهل أو صعب)
        function startQuiz(mode) {
            currentMode = mode;
            questions = [];
            answeredOperations = [];
            for (let i = 0; i <= 10; i++) {
                questions.push({ multiplier: i, answer: currentNumber * i });
            }
            if (mode === 'hard') {
                questions = questions.sort(() => Math.random() - 0.5);
            }
            currentQuestion = 0;
            score = 0;
            mistakes = [];
            document.getElementById('optionsMenu').style.display = 'none';
            document.getElementById('quizView').style.display = 'block';
            document.getElementById('quizTitle').textContent = `سلسلة ${mode === 'easy' ? 'سهلة' : 'صعبة'} للرقم ${currentNumber}`;
            document.getElementById('completedOperations').innerHTML = '';
            showQuestion();
        }

        // عرض سؤال جديد
        function showQuestion() {
            const backButton = document.getElementById('backFromQuizBtn');
            if (backButton) {
                backButton.style.display = 'inline-block';
            }
            if (currentQuestion >= questions.length) {
                showResults();
                return;
            }
            const q = questions[currentQuestion];
            document.getElementById('question').textContent = `${currentNumber} × ${q.multiplier} = ?`;
            const answersDiv = document.getElementById('answers');
            answersDiv.innerHTML = '';
            const correctAnswer = q.answer;
            const options = [correctAnswer];
            while (options.length < 3) {
                const randomAnswer = Math.max(0, correctAnswer + Math.floor(Math.random() * 10));
                if (randomAnswer !== correctAnswer && !options.includes(randomAnswer)) {
                    options.push(randomAnswer);
                }
            }
            options.sort(() => Math.random() - 0.5);
            options.forEach(opt => {
                const btn = document.createElement('button');
                btn.className = 'answer-btn';
                btn.textContent = opt;
                btn.onclick = () => checkAnswer(opt, correctAnswer, btn);
                answersDiv.appendChild(btn);
            });
            document.getElementById('feedback').textContent = '';
        }

        // التحقق من الإجابة المختارة
        function checkAnswer(selected, correct, btn) {
            const q = questions[currentQuestion];
            const completed = document.getElementById('completedOperations');
            const feedback = document.getElementById('feedback');
            const backButton = document.getElementById('backFromQuizBtn');
            if (backButton) {
                backButton.style.display = 'none';
            }
            const allButtons = document.querySelectorAll('.answer-btn');
            allButtons.forEach(b => b.disabled = true);
            const operation = {
                multiplier: q.multiplier,
                correctAnswer: correct,
                selectedAnswer: selected,
                isCorrect: selected === correct
            };
            answeredOperations.push(operation);
            if (selected === correct) {
                score++;
                btn.classList.add('correct-answer');
                feedback.textContent = ['أحسنت!', 'ممتاز!', 'رائع!'][Math.floor(Math.random() * 3)];
                feedback.className = 'feedback correct';
                completed.innerHTML += `<div class="operation-item correct">${currentNumber} × ${q.multiplier} = ${correct}</div>`;
                playCorrectSound();
            } else {
                btn.classList.add('wrong-answer');
                const correctButton = Array.from(allButtons).find(b => Number(b.textContent) === correct);
                if (correctButton) {
                    correctButton.classList.add('correct-answer');
                }
                feedback.textContent = `خطأ! الإجابة الصحيحة هي ${correct}`;
                feedback.className = 'feedback wrong';
                completed.innerHTML += `<div class="operation-item wrong">${currentNumber} × ${q.multiplier} = ${correct}</div>`;
                mistakes.push(`${currentNumber} × ${q.multiplier} = ${correct}`);
                playWrongSound();
                allErrors.push({
                    mistake: `${currentNumber} × ${q.multiplier} = ${correct}`,
                    date: new Date().toLocaleString('en-US', { numberingSystem: 'latn' })
                });
                localStorage.setItem('multiplicationErrors', JSON.stringify(allErrors));
                updateErrorsBadge();
            }
            completed.scrollTop = completed.scrollHeight;
            const lastCompleted = completed.lastElementChild;
            if (lastCompleted) {
                lastCompleted.scrollIntoView({ behavior: 'smooth', block: 'end' });
            }
            currentQuestion++;
            setTimeout(showQuestion, 1000);
        }

        // عرض نتائج الاختبار
        function showResults() {
            document.getElementById('quizView').style.display = 'none';
            document.getElementById('resultView').style.display = 'block';
            document.getElementById('resultSummary').textContent = `لقد أجبت على ${score} من ${questions.length} بشكل صحيح!`;
            const mistakesDiv = document.getElementById('mistakes');
            const allOperationsDiv = document.getElementById('allOperations');
            mistakesDiv.innerHTML = '<h3>الأخطاء:</h3>';
            if (mistakes.length === 0) {
                mistakesDiv.innerHTML += '<div class="history-item correct">لم ترتكب أي أخطاء! رائع!</div>';
            } else {
                mistakes.forEach(m => {
                    mistakesDiv.innerHTML += `<div class="history-item wrong">${m}</div>`;
                });
            }
            allOperationsDiv.innerHTML = '<h3>جميع العمليات في السلسلة:</h3>';
            answeredOperations.forEach(op => {
                const statusClass = op.isCorrect ? 'correct' : 'wrong';
                const statusText = op.isCorrect ? 'صحيح' : `خطأ (اختيارك: ${op.selectedAnswer})`;
                allOperationsDiv.innerHTML += `<div class="history-item ${statusClass}">${currentNumber} × ${op.multiplier} = ${op.correctAnswer} — ${statusText}</div>`;
            });
            playCelebrationSound();
        }

        // عرض الأخطاء السابقة مع التواريخ
        function updateErrorsBadge() {
            const badge = document.getElementById('errorsBadge');
            if (!badge) return;
            const count = allErrors.length;
            badge.textContent = count > 0 ? count : '';
        }

        function showErrors() {
            document.getElementById('mainMenu').style.display = 'none';
            document.getElementById('errorsView').style.display = 'block';
            const errorsList = document.getElementById('errorsList');
            errorsList.innerHTML = '';
            if (allErrors.length === 0) {
                errorsList.innerHTML = '<p>لا توجد أخطاء مسجلة.</p>';
            } else {
                allErrors.forEach(err => {
                    errorsList.innerHTML += `<div class="history-item wrong"><span class="error-item-badge"></span>${err.mistake} - ${err.date}</div>`;
                });
            }
        }

        // الرجوع إلى القائمة الرئيسية
        function backToMain() {
            document.getElementById('optionsMenu').style.display = 'none';
            document.getElementById('tableView').style.display = 'none';
            document.getElementById('quizView').style.display = 'none';
            document.getElementById('resultView').style.display = 'none';
            document.getElementById('errorsView').style.display = 'none';
            document.getElementById('mainMenu').style.display = 'block';
        }

        // الرجوع إلى قائمة الخيارات
        function backToOptions() {
            document.getElementById('tableView').style.display = 'none';
            document.getElementById('quizView').style.display = 'none';
            document.getElementById('resultView').style.display = 'none';
            document.getElementById('optionsMenu').style.display = 'block';
        }

        // تهيئة اللعبة عند تحميل الصفحة
        function initWelcomeOverlay() {
            const overlay = document.getElementById('welcomeOverlay');
            const voirBtn = document.getElementById('voirPlusBtn');
            const details = document.getElementById('contactDetails');
            const startBtn = document.getElementById('startGameBtn');
            const carouselTrack = document.getElementById('carouselTrack');
            const carouselViewport = document.getElementById('carouselViewport');

            const resolveAssetUrl = (relativePath) => new URL(relativePath, document.baseURI).href;

            const slides = [
                resolveAssetUrl('./images/photo1.png'),
                resolveAssetUrl('./images/photo2.png'),
                resolveAssetUrl('./images/photo3.png')
            ];

            carouselTrack.innerHTML = '';

            slides.forEach(src => {
                const slide = document.createElement('div');
                slide.className = 'carousel-slide';
                const img = document.createElement('img');
                img.src = src;
                img.alt = 'Image du carousel';
                img.loading = 'eager';
                slide.appendChild(img);
                slide.setAttribute('data-src', src);
                carouselTrack.appendChild(slide);
            });

            // Add Navigation Arrows
            const prevArrow = document.createElement('button');
            prevArrow.className = 'carousel-arrow prev';
            prevArrow.innerHTML = '‹';
            prevArrow.type = 'button';
            prevArrow.ariaLabel = 'الصورة السابقة';
            carouselViewport.appendChild(prevArrow);

            const nextArrow = document.createElement('button');
            nextArrow.className = 'carousel-arrow next';
            nextArrow.innerHTML = '›';
            nextArrow.type = 'button';
            nextArrow.ariaLabel = 'الصورة التالية';
            carouselViewport.appendChild(nextArrow);

            // Add Dots Container and Dots
            const dotsContainer = document.createElement('div');
            dotsContainer.className = 'carousel-dots';
            carouselViewport.appendChild(dotsContainer);

            const dots = [];
            slides.forEach((_, index) => {
                const dot = document.createElement('button');
                dot.className = 'carousel-dot' + (index === 0 ? ' active' : '');
                dot.type = 'button';
                dot.ariaLabel = `الانتقال إلى الصورة ${index + 1}`;
                dotsContainer.appendChild(dot);
                dots.push(dot);

                dot.addEventListener('click', () => {
                    goToSlide(index);
                    resetAutoplay();
                });
            });

            let currentIndex = 0;
            let isDragging = false;
            let startX = 0;
            let currentTranslate = 0;
            let dragThreshold = 50;
            let isMoving = false;
            let startTranslate = 0;

            const updateDots = () => {
                const activeDotIndex = currentIndex % slides.length;
                dots.forEach((dot, idx) => {
                    dot.classList.toggle('active', idx === activeDotIndex);
                });
            };

            const setPositionByIndex = () => {
                const viewportWidth = carouselViewport.clientWidth || 300;
                currentTranslate = -currentIndex * viewportWidth;
                carouselTrack.style.transform = `translateX(${currentTranslate}px)`;
                updateDots();
            };

            const resetCarouselPosition = () => {
                currentIndex = 0;
                carouselTrack.style.transition = 'none';
                carouselTrack.style.transform = 'translateX(0px)';
                updateDots();
            };

            const goToSlide = (index) => {
                currentIndex = (index + slides.length) % slides.length;
                carouselTrack.style.transition = 'transform 0.3s cubic-bezier(0.25, 0.46, 0.45, 0.94)';
                requestAnimationFrame(() => {
                    const viewportWidth = carouselViewport.clientWidth || 300;
                    currentTranslate = -currentIndex * viewportWidth;
                    carouselTrack.style.transform = `translateX(${currentTranslate}px)`;
                    updateDots();
                });
            };

            // Event Listeners for Arrows
            prevArrow.addEventListener('click', () => {
                goToSlide(currentIndex - 1);
                resetAutoplay();
            });

            nextArrow.addEventListener('click', () => {
                goToSlide(currentIndex + 1);
                resetAutoplay();
            });

            // Adjust position on resize
            window.addEventListener('resize', () => {
                carouselTrack.style.transition = 'none';
                setPositionByIndex();
            });

            carouselTrack.addEventListener('transitionend', () => {
                carouselTrack.style.transition = 'none';
            });

            // Pointer Events for Drag / Swipe (touch and mouse)
            carouselViewport.addEventListener('pointerdown', (e) => {
                isDragging = true;
                isMoving = false;
                startX = e.clientX;
                startTranslate = -currentIndex * carouselViewport.clientWidth;
                carouselTrack.style.transition = 'none';
                carouselViewport.classList.add('dragging');
                carouselViewport.setPointerCapture(e.pointerId);
                pauseAutoplay();
            });

            carouselViewport.addEventListener('pointermove', (e) => {
                if (!isDragging) return;
                const currentX = e.clientX;
                const diffX = currentX - startX;
                
                if (Math.abs(diffX) > 5) {
                    isMoving = true;
                }

                currentTranslate = startTranslate + diffX;
                carouselTrack.style.transform = `translateX(${currentTranslate}px)`;
            });

            const handlePointerUp = (e) => {
                if (!isDragging) return;
                isDragging = false;
                carouselViewport.classList.remove('dragging');
                try {
                    carouselViewport.releasePointerCapture(e.pointerId);
                } catch (_) {}

                const movedBy = e.clientX - startX;
                carouselTrack.style.transition = 'transform 0.3s cubic-bezier(0.25, 0.46, 0.45, 0.94)';

                if (isMoving && Math.abs(movedBy) > dragThreshold) {
                    if (movedBy < 0) {
                        // Dragged left, go to next slide
                        goToSlide(currentIndex + 1);
                    } else {
                        // Dragged right, go to previous slide
                        goToSlide(currentIndex - 1);
                    }
                } else {
                    // Snap back to original slide
                    goToSlide(currentIndex);
                }
                
                resumeAutoplay();
            };

            carouselViewport.addEventListener('pointerup', handlePointerUp);
            carouselViewport.addEventListener('pointercancel', handlePointerUp);

            // Image Zoom Modal implementation
            const imageModal = document.getElementById('imageModal');
            const modalImage = document.getElementById('modalImage');
            const modalClose = imageModal.querySelector('.image-modal-close');

            const openImageModal = (src) => {
                modalImage.src = src;
                imageModal.classList.add('show');
                imageModal.setAttribute('aria-hidden', 'false');
            };
            
            const closeImageModal = () => {
                imageModal.classList.remove('show');
                imageModal.setAttribute('aria-hidden', 'true');
                setTimeout(() => {
                    if (!imageModal.classList.contains('show')) {
                        modalImage.src = '';
                    }
                }, 300);
            };

            modalClose.addEventListener('click', closeImageModal);
            imageModal.addEventListener('click', (event) => {
                if (event.target === imageModal) {
                    closeImageModal();
                }
            });
            document.addEventListener('keydown', (event) => {
                if (event.key === 'Escape' && imageModal.classList.contains('show')) {
                    closeImageModal();
                }
            });

            // Set up click on slides to trigger lightbox
            carouselTrack.querySelectorAll('.carousel-slide').forEach(slide => {
                slide.addEventListener('click', (e) => {
                    // Only open if the user clicked (was not dragging)
                    if (!isMoving) {
                        const src = slide.getAttribute('data-src');
                        openImageModal(src);
                    }
                });
            });

            // Autoplay Functionality
            let autoplayInterval = null;

            const startAutoplay = () => {
                if (autoplayInterval) return;
                autoplayInterval = setInterval(() => {
                    goToSlide(currentIndex + 1);
                }, 3500);
            };

            const pauseAutoplay = () => {
                if (autoplayInterval) {
                    clearInterval(autoplayInterval);
                    autoplayInterval = null;
                }
            };

            const resumeAutoplay = () => {
                startAutoplay();
            };

            const resetAutoplay = () => {
                pauseAutoplay();
                resumeAutoplay();
            };

            // Pause on hover
            carouselViewport.addEventListener('mouseenter', pauseAutoplay);
            carouselViewport.addEventListener('mouseleave', resumeAutoplay);

            // Initialize position and start autoplay
            resetCarouselPosition();
            requestAnimationFrame(() => {
                resetCarouselPosition();
            });
            startAutoplay();

            // Toggle Voir Plus
            voirBtn.addEventListener('click', () => {
                const isVisible = details.style.display === 'block';
                details.style.display = isVisible ? 'none' : 'block';
                voirBtn.textContent = isVisible ? 'عرض المزيد' : 'عرض أقل';
            });

            // Start game button
            startBtn.addEventListener('click', () => {
                overlay.style.display = 'none';
                backToMain();
            });
        }

        initWelcomeOverlay();
        initNumberGrid();
        initBackgroundMusic();
        document.getElementById('showErrorsBtn').onclick = showErrors;
        updateErrorsBadge();
