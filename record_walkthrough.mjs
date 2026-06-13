import puppeteer from 'puppeteer-core';
import fs from 'fs';
import path from 'path';

async function run() {
    const brainDir = "/Users/balbaasaur/.gemini/antigravity-ide/brain/d69807e3-4a91-495c-9f6f-3c324bb4b0f8";
    const framesDir = path.join(brainDir, "scratch", "hd_frames");
    
    if (fs.existsSync(framesDir)) {
        fs.rmSync(framesDir, { recursive: true, force: true });
    }
    fs.mkdirSync(framesDir, { recursive: true });

    console.log("Launching headless Google Chrome...");
    const browser = await puppeteer.launch({
        executablePath: '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
        headless: "new",
        args: ['--no-sandbox', '--disable-setuid-sandbox']
    });

    const page = await browser.newPage();
    await page.setViewport({ width: 1920, height: 1080 });

    console.log("Opening dashboard walkthrough page...");
    await page.goto('file:///Users/balbaasaur/Documents/D2C%20Google%20Sheet%20Dashboard%20Builder/index.html?walkthrough=true');

    console.log("Starting 100-second Full HD frame capture loop at 10 FPS...");
    let frameCount = 0;
    const startTime = Date.now();
    const duration = 100 * 1000; // 100 seconds
    const interval = 100; // 100ms per frame

    while (Date.now() - startTime < duration) {
        const frameStart = Date.now();
        frameCount++;
        const filename = path.join(framesDir, `frame_${String(frameCount).padStart(5, '0')}.png`);
        
        await page.screenshot({ path: filename, type: 'png' });
        
        const elapsed = Date.now() - frameStart;
        const sleepTime = Math.max(0, interval - elapsed);
        
        if (frameCount % 50 === 0) {
            const pct = Math.round(((Date.now() - startTime) / duration) * 100);
            console.log(`Captured frame #${frameCount} (${pct}% complete)...`);
        }
        
        if (sleepTime > 0) {
            await new Promise(r => setTimeout(r, sleepTime));
        }
    }

    console.log(`Full HD walkthrough frame capture completed. Total frames: ${frameCount}`);
    await browser.close();
}

run().catch(console.error);
