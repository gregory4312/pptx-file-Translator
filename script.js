var uploadedFile = null;
var originalZip = null;
var slideTexts = [];

function handleFileSelect(file) {
    if (!file) return;
    
    if (!file.name.toLowerCase().endsWith('.pptx')) {
        showStatus('Please select a .pptx file', 'error');
        return;
    }

    uploadedFile = file;
    
    var fileInfo = document.getElementById('fileInfo');
    if (fileInfo) {
        fileInfo.innerHTML = '<strong>File:</strong> ' + file.name + '<br>' +
            '<strong>Size:</strong> ' + (file.size / 1024 / 1024).toFixed(2) + ' MB<br>' +
            '<strong>Status:</strong> Ready for translation';
        fileInfo.style.display = 'block';
    }
    
    var translateBtn = document.getElementById('translateBtn');
    if (translateBtn) {
        translateBtn.disabled = false;
    }
    
    showStatus('File uploaded successfully! Select languages and click translate.', 'success');
}

function swapLanguages() {
    var fromLang = document.getElementById('fromLang');
    var toLang = document.getElementById('toLang');
    
    if (fromLang.value === 'auto') {
        showStatus('Cannot swap from auto-detect', 'error');
        return;
    }
    
    var temp = fromLang.value;
    fromLang.value = toLang.value;
    toLang.value = temp;
}

function startTranslation() {
    if (!uploadedFile) {
        showStatus('Please upload a PowerPoint file first', 'error');
        return;
    }

    var fromLang = document.getElementById('fromLang').value;
    var toLang = document.getElementById('toLang').value;

    if (fromLang === toLang) {
        showStatus('Source and target languages cannot be the same', 'error');
        return;
    }

    var translateBtn = document.getElementById('translateBtn');
    if (translateBtn) {
        translateBtn.disabled = true;
        translateBtn.classList.add('processing');
    }
    
    showProgress(0, 'Starting...', 'Preparing to load your PowerPoint file');
    showStatus('Loading PowerPoint file...', 'info');

    loadPowerPoint(uploadedFile)
        .then(function() {
            showProgress(20, 'File Loaded', 'Found ' + slideTexts.length + ' slides with text');
            var totalTexts = slideTexts.reduce((sum, slide) => sum + slide.textElements.length, 0);

            if (totalTexts === 0) {
                showStatus('No text found in the presentation', 'error');
                resetUI();
                return;
            }

            showStatus('Translating text...', 'info');
            showProgress(40, 'Translating...', 'Translating ' + totalTexts + ' text elements');
            return translateAllText(fromLang, toLang);
        })
        .then(function() {
            showProgress(80, 'Creating File...', 'Building your translated PowerPoint presentation');
            showStatus('Creating translated presentation...', 'info');
            return createTranslatedPowerPoint();
        })
        .then(function() {
            showProgress(100, 'Complete!', 'Your translated presentation is ready for download');
            showStatus('Translation completed! Download should start automatically.', 'success');
        })
        .catch(function(error) {
            console.error('Translation error:', error);
            showStatus('Translation failed: ' + error.message, 'error');
        })
        .finally(function() {
            resetUI();
        });
}

function loadPowerPoint(file) {
    return new Promise(function(resolve, reject) {
        var zip = new JSZip();
        zip.loadAsync(file)
            .then(function(zipData) {
                originalZip = zipData;
                slideTexts = [];

                var slideFiles = Object.keys(originalZip.files).filter(function(name) {
                    return name.startsWith('ppt/slides/slide') && name.endsWith('.xml');
                }).sort();

                var promises = slideFiles.map(function(slideFile) {
                    return originalZip.files[slideFile].async('text')
                        .then(function(slideContent) {
                            if (typeof DOMParser === 'undefined') {
                                showStatus('Your browser does not support DOMParser. Translation may fail for complex files.', 'error');
                                reject(new Error('DOMParser not supported'));
                                return;
                            }
                            
                            var parser = new DOMParser();
                            var xmlDoc = parser.parseFromString(slideContent, "text/xml");
                            var textElements = extractTextElementsFromXML(xmlDoc);

                            if (textElements.length > 0) {
                                slideTexts.push({
                                    slideFile: slideFile,
                                    xmlDoc: xmlDoc, // Store the entire XML document object
                                    textElements: textElements
                                });
                            }
                        });
                });

                return Promise.all(promises);
            })
            .then(resolve)
            .catch(reject);
    });
}

function extractTextElementsFromXML(xmlDoc) {
    var textElements = [];
    
    // Handle namespace for PowerPoint XML
    var textNodes = xmlDoc.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/main", "t");
    
    for (var i = 0; i < textNodes.length; i++) {
        var textNode = textNodes[i];
        var textContent = textNode.textContent ? textNode.textContent.trim() : '';
        
        if (textContent) {
            textElements.push({
                originalText: textContent,
                translatedText: null,
                xmlNode: textNode
            });
        }
    }
    
    return textElements;
}

function createTranslatedPowerPoint() {
    return new Promise(function(resolve, reject) {
        if (!originalZip) {
            reject(new Error("No PowerPoint file loaded"));
            return;
        }

        var newZip = new JSZip();
        var copyPromises = [];
        
        Object.keys(originalZip.files).forEach(function(filename) {
            var file = originalZip.files[filename];
            
            if (file.dir) {
                newZip.folder(filename);
            } else {
                copyPromises.push(
                    file.async('uint8array').then(function(content) {
                        return { filename: filename, content: content };
                    })
                );
            }
        });

        Promise.all(copyPromises)
            .then(function(files) {
                files.forEach(function(fileData) {
                    newZip.file(fileData.filename, fileData.content);
                });

                var slidePromises = slideTexts.map(function(slideData) {
                    var modifiedXmlDoc = slideData.xmlDoc;
                    
                    slideData.textElements.forEach(function(textElement) {
                        if (textElement.translatedText && textElement.translatedText !== textElement.originalText) {
                            textElement.xmlNode.textContent = textElement.translatedText;
                        }
                    });
                    
                    var serializer = new XMLSerializer();
                    var modifiedXml = serializer.serializeToString(modifiedXmlDoc);
                    
                    newZip.file(slideData.slideFile, modifiedXml);
                });

                return Promise.all(slidePromises);
            })
            .then(function() {
                return newZip.generateAsync({
                    type: 'blob',
                    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
                    compression: 'DEFLATE',
                    compressionOptions: { level: 6 }
                });
            })
            .then(function(blob) {
                var url = URL.createObjectURL(blob);
                var a = document.createElement('a');
                a.href = url;
                a.download = 'translated_' + (uploadedFile ? uploadedFile.name : 'presentation.pptx');
                document.body.appendChild(a);
                a.click();
                setTimeout(function() {
                    document.body.removeChild(a);
                    URL.revokeObjectURL(url);
                    resolve();
                }, 100);
            })
            .catch(function(error) {
                console.error('Error creating PPTX:', error);
                reject(new Error('Failed to create translated presentation. Please try a simpler file or different languages.'));
            });
    });
}

function translateAllText(fromLang, toLang) {
    return new Promise(function(resolve, reject) {
        var allTextElements = [];
        slideTexts.forEach(slide => {
            allTextElements.push(...slide.textElements);
        });

        var processedCount = 0;
        var totalCount = allTextElements.length;

        if (totalCount === 0) {
            resolve();
            return;
        }

        function translateNext() {
            if (processedCount >= totalCount) {
                resolve();
                return;
            }

            var textElement = allTextElements[processedCount];
            
            translateWithGoogleFree(textElement.originalText, fromLang, toLang)
                .then(function(translatedText) {
                    textElement.translatedText = translatedText;
                    processedCount++;
                    
                    var progressPercent = 40 + (processedCount / totalCount) * 40;
                    var remaining = totalCount - processedCount;
                    showProgress(
                        progressPercent, 
                        'Translating...', 
                        'Translated ' + processedCount + ' of ' + totalCount + ' text elements (' + remaining + ' remaining)'
                    );
                    
                    setTimeout(translateNext, 300);
                })
                .catch(function(error) {
                    console.error('Error translating:', textElement.originalText, error);
                    textElement.translatedText = textElement.originalText;
                    processedCount++;
                    setTimeout(translateNext, 300);
                });
        }

        translateNext();
    });
}

function translateWithGoogleFree(text, fromLang, toLang) {
    return new Promise(function(resolve, reject) {
        var encodedText = encodeURIComponent(text);
        var sourceLang = fromLang === 'auto' ? 'auto' : fromLang;
        var url = 'https://translate.googleapis.com/translate_a/single?client=gtx&sl=' + 
                  sourceLang + '&tl=' + toLang + '&dt=t&q=' + encodedText;
        
        fetch(url)
            .then(function(response) {
                if (!response.ok) throw new Error('Translation failed');
                return response.json();
            })
            .then(function(data) {
                if (data && data[0] && data[0][0] && data[0][0][0]) {
                    resolve(data[0][0][0]);
                } else {
                    throw new Error('Invalid response');
                }
            })
            .catch(function() {
                fallbackTranslation(text, fromLang, toLang)
                    .then(resolve)
                    .catch(function() {
                        resolve(text); // Return original text if translation fails
                    });
            });
    });
}

function fallbackTranslation(text, fromLang, toLang) {
    return new Promise(function(resolve, reject) {
        var encodedText = encodeURIComponent(text);
        var langPair = (fromLang === 'auto' ? 'en' : fromLang) + '|' + toLang;
        var url = 'https://api.mymemory.translated.net/get?q=' + encodedText + '&langpair=' + langPair;
        
        fetch(url)
            .then(function(response) {
                if (!response.ok) throw new Error('Fallback failed');
                return response.json();
            })
            .then(function(data) {
                if (data && data.responseData && data.responseData.translatedText) {
                    resolve(data.responseData.translatedText);
                } else {
                    resolve(text);
                }
            })
            .catch(function() {
                resolve(text);
            });
    });
}

function showStatus(message, type) {
    var statusEl = document.getElementById('status');
    if (statusEl) {
        statusEl.textContent = message;
        statusEl.className = 'status ' + type;
        statusEl.style.display = 'block';
    }
}

function showProgress(percent, text, details) {
    var progressEl = document.getElementById('progress');
    var progressFillEl = document.getElementById('progressFill');
    var progressTextEl = document.getElementById('progressText');
    var progressPercentageEl = document.getElementById('progressPercentage');
    var progressDetailsEl = document.getElementById('progressDetails');
    
    if (progressEl && progressFillEl) {
        progressEl.style.display = 'block';
        progressFillEl.style.width = percent + '%';
        
        if (progressPercentageEl) {
            progressPercentageEl.textContent = Math.round(percent) + '%';
        }
        
        if (progressTextEl && text) {
            progressTextEl.textContent = text;
        }
        
        if (progressDetailsEl && details) {
            progressDetailsEl.textContent = details;
        }
        
        if (percent >= 100) {
            setTimeout(function() {
                if (progressEl) {
                    progressEl.style.display = 'none';
                    progressFillEl.style.width = '0%';
                    if (progressPercentageEl) progressPercentageEl.textContent = '0%';
                }
            }, 3000);
        }
    }
}

function resetUI() {
    var translateBtn = document.getElementById('translateBtn');
    if (translateBtn) {
        translateBtn.disabled = !uploadedFile;
        translateBtn.classList.remove('processing');
    }
}

window.addEventListener('load', function() {
    var fileUpload = document.getElementById('fileUpload');
    var fileInput = document.getElementById('fileInput');
    
    if (fileUpload && fileInput) {
        fileUpload.addEventListener('dragover', function(e) {
            e.preventDefault();
            fileUpload.classList.add('dragover');
        });

        fileUpload.addEventListener('dragleave', function() {
            fileUpload.classList.remove('dragover');
        });

        fileUpload.addEventListener('drop', function(e) {
            e.preventDefault();
            fileUpload.classList.remove('dragover');
            var files = e.dataTransfer.files;
            if (files.length > 0) {
                handleFileSelect(files[0]);
            }
        });
    }
});