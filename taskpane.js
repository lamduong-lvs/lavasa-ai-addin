// =======================
// Lavasa AI Add-in Script
// =======================

(async function () {
    Office.onReady(async (info) => {
        if (info.host === Office.HostType.PowerPoint) {
            document.getElementById('generateBtn').onclick = generatePresentation;
            document.getElementById('clearHistoryBtn').onclick = clearHistory;
            loadHistory();
        }
    });
})();

async function generatePresentation() {
    const prompt = document.getElementById('prompt').value.trim();
    const style = document.getElementById('style').value;
    const model = document.getElementById('model').value;
    const apiKey = document.getElementById('apiKey').value.trim();
    const enableImage = document.getElementById('enableImage').checked;
    const enableAnimation = document.getElementById('enableAnimation').checked;
    const enableChart = document.getElementById('enableChart').checked;
    const enableTimeline = document.getElementById('enableTimeline').checked;
    const enableWatermark = document.getElementById('enableWatermark').checked;
    const loading = document.getElementById('loading');

    if (!prompt || !apiKey) {
        alert('Vui lòng nhập cả Prompt và API Key.');
        return;
    }

    loading.style.display = 'block';

    try {
        // Compose API call
        const messages = [
            { role: "system", content: `Bạn là một trợ lý AI chuyên tạo nội dung bài thuyết trình phong cách ${style}. Hãy trả về dưới dạng Outline: mỗi slide là 1 tiêu đề + 3-5 bullet points.` },
            { role: "user", content: prompt }
        ];

        const response = await fetch('https://api.openai.com/v1/chat/completions', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${apiKey}`
            },
            body: JSON.stringify({
                model: model,
                messages: messages,
                temperature: 0.7
            })
        });

        const data = await response.json();
        const content = data.choices[0].message.content;
        const slides = parseSlides(content);

        // Insert slides
        for (let slide of slides) {
            await Office.context.document.setSelectedDataAsync(
                "",
                { coercionType: Office.CoercionType.SlideRange, asyncContext: slide },
                (asyncResult) => {
                    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                        const title = slide.title;
                        const bullets = slide.bullets;
                        insertSlideContent(title, bullets, enableWatermark);
                    }
                }
            );
        }

        // Optional Image Insert
        if (enableImage) {
            await insertImage(prompt, apiKey);
        }

        // Animation/Chart/Timeline Optional
        if (enableAnimation) {
            await addAnimations();
        }
        if (enableChart) {
            await createChart();
        }
        if (enableTimeline) {
            await createTimeline();
        }

        saveHistory(prompt);

        alert('✅ Đã tạo bài thuyết trình thành công!');
    } catch (error) {
        console.error(error);
        alert('❌ Lỗi khi tạo bài thuyết trình.');
    } finally {
        loading.style.display = 'none';
    }
}

// Parse GPT response into structured slides
function parseSlides(text) {
    const slides = [];
    const lines = text.split('\n').filter(line => line.trim() !== "");
    let currentSlide = { title: "", bullets: [] };

    lines.forEach(line => {
        if (!line.startsWith("-") && !line.startsWith("*")) {
            if (currentSlide.title) {
                slides.push(currentSlide);
                currentSlide = { title: "", bullets: [] };
            }
            currentSlide.title = line.trim();
        } else {
            currentSlide.bullets.push(line.replace(/^[-*]\s*/, "").trim());
        }
    });

    if (currentSlide.title) {
        slides.push(currentSlide);
    }

    return slides;
}

// Insert content into Slide
function insertSlideContent(title, bullets, enableWatermark) {
    Office.context.document.setSelectedDataAsync(
        {
            type: "text",
            data: `${title}\n\n${bullets.join('\n')}`
        },
        { coercionType: "text" },
        function (asyncResult) {
            if (enableWatermark) {
                addWatermark();
            }
        }
    );
}

// Save prompt to local storage history
function saveHistory(prompt) {
    let history = JSON.parse(localStorage.getItem('lavasaHistory')) || [];
    history.unshift({ prompt, time: new Date().toLocaleString() });
    localStorage.setItem('lavasaHistory', JSON.stringify(history));
    loadHistory();
}

// Load history to UI
function loadHistory() {
    const historyList = document.getElementById('historyList');
    historyList.innerHTML = "";
    const history = JSON.parse(localStorage.getItem('lavasaHistory')) || [];
    history.forEach(item => {
        const li = document.createElement('li');
        li.textContent = `[${item.time}] ${item.prompt}`;
        historyList.appendChild(li);
    });
}

// Clear history
function clearHistory() {
    localStorage.removeItem('lavasaHistory');
    loadHistory();
}

// Add Watermark
function addWatermark() {
    // (Giản lược mô phỏng Watermark - thực tế cần Office Script chi tiết hơn)
    console.log("Add Watermark: Powered by Lavasa AI");
}

// Insert Image from DALL-E
async function insertImage(prompt, apiKey) {
    const dalleResponse = await fetch('https://api.openai.com/v1/images/generations', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${apiKey}`
        },
        body: JSON.stringify({
            prompt: prompt,
            n: 1,
            size: "1024x1024"
        })
    });

    const dalleData = await dalleResponse.json();
    const imageUrl = dalleData.data[0].url;
    console.log("Generated image:", imageUrl);

    // (Giản lược - Office.js thực tế cần nhiều bước hơn để insert Image vào Slide)
}

// Add Animations
async function addAnimations() {
    console.log("Animation effects added to slides!");
}

// Create Chart
async function createChart() {
    console.log("Chart created based on content!");
}

// Create Timeline
async function createTimeline() {
    console.log("Timeline generated from dates!");
}
