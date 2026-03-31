// pptx-generator.js
export async function generateNativePptx(slides, downloadName = "Finam_Presentation.pptx") {
    if (!window.PptxGenJS) {
        throw new Error("PptxGenJS не загружен");
    }
    const pres = new PptxGenJS();
    pres.layout = "LAYOUT_16x9"; // 10 x 5.625 inches

    // Хелперы для перевода пикселей (1920x1080) в дюймы (10 x 5.625)
    const pX = (px) => px / 192;
    const pY = (px) => px / 192;
    const fZ = (px) => px * 0.75; // px to pt

    const bg1 = "/assets/cd50dbb4125036734ac5a2db58ea9629ef8c0469.png";
    const bg2 = "/assets/6361b1c3ce7c16cafa967d0a2a49e6cd08cff83d.png";
    const bgDark = "/assets/dc8112d12f5ea2931c46b016d4284646f463d47e.png";
    const bgDefault = "/assets/91c825f6107036a4b04b5de1798624c0250fa53c.png";
    const logoMain = "/assets/4cce4941abc900d8328cc7c84a332b8d0251ec15.png";
    const logoFooter = "/assets/1fcb1065a69e3b98f5d3209cc5c4a157837f949e.png";
    const avatar = "/assets/20af95fd83ad18f7f8f3cfd9a46040bd59dfada7.png";
    const placeholder1 = "/assets/cf020cf52bfc734d596e06e6ed617b2f7fbe34d4.png";
    const placeholder2 = "/assets/e9cb1099bc29ec9dc62e7b665222b3619e294c9b.png";

    const commonTitle = (s, text) => {
        if (text) s.addText(text, { x: pX(88), y: pY(61), w: pX(1744), h: pY(60), fontSize: fZ(48), color: "FFFFFF", fontFace: "Inter", valign: "top" });
    };
    const addFooter = (s) => {
        s.addImage({ path: logoFooter, x: pX(1606), y: pY(953), w: pX(225), h: pY(62) });
    };

    for (const data of slides) {
        const slide = pres.addSlide();
        const f = data.fields;
        const tid = data.templateId;

        if (tid === "template-1") {
            slide.addImage({ path: bg1, x: 0, y: 0, w: '100%', h: '100%' });
            slide.addImage({ path: logoMain, x: pX(88), y: pY(61), w: pX(301.62), h: pY(83) });
            slide.addImage({ path: avatar, x: pX(88), y: pY(896), w: pX(121), h: pY(121) });
            slide.addText(f.title, { x: pX(88), y: pY(300), w: pX(1700), h: pY(120), fontSize: fZ(122), color: "FFFFFF", fontFace: "Inter", bold: true, valign: "bottom" });
            slide.addText(f.subtitle, { x: pX(88), y: pY(440), w: pX(1700), h: pY(60), fontSize: fZ(48), color: "FFFFFF", fontFace: "Inter" });
            slide.addText(f.name || "", { x: pX(238), y: pY(910), w: pX(600), h: pY(30), fontSize: fZ(28), color: "FFFFFF", fontFace: "Inter", bold: true });
            slide.addText(f.position || "", { x: pX(238), y: pY(950), w: pX(600), h: pY(30), fontSize: fZ(22), color: "A6A6A6", fontFace: "Inter" });
        }
        else if (tid === "template-2") {
            slide.addImage({ path: bg2, x: 0, y: 0, w: '100%', h: '100%' });
            slide.addImage({ path: logoMain, x: pX(88), y: pY(61), w: pX(301.62), h: pY(83) });
            slide.addImage({ path: avatar, x: pX(88), y: pY(896), w: pX(121), h: pY(121) });
            slide.addText(f.title, { x: pX(88), y: pY(300), w: pX(1700), h: pY(120), fontSize: fZ(122), color: "FFFFFF", fontFace: "Inter", bold: true, valign: "bottom" });
            slide.addText(f.subtitle, { x: pX(88), y: pY(440), w: pX(1700), h: pY(60), fontSize: fZ(48), color: "FFFFFF", fontFace: "Inter" });
            slide.addText(f.description, { x: pX(88), y: pY(540), w: pX(1332), h: pY(160), fontSize: fZ(22), color: "A6A6A6", fontFace: "Inter" });
            slide.addText(f.name || "", { x: pX(238), y: pY(910), w: pX(600), h: pY(30), fontSize: fZ(28), color: "FFFFFF", fontFace: "Inter", bold: true });
            slide.addText(f.position || "", { x: pX(238), y: pY(950), w: pX(600), h: pY(30), fontSize: fZ(22), color: "A6A6A6", fontFace: "Inter" });
        }
        else if (tid === "template-3" || tid === "template-5") {
            slide.addImage({ path: bgDark, x: 0, y: 0, w: '100%', h: '100%' });
            slide.addImage({ path: logoMain, x: pX(88), y: pY(61), w: pX(301.62), h: pY(83) });
            slide.addText(f.divider_text, { x: pX(88), y: pY(350), w: pX(1700), h: pY(400), fontSize: fZ(90), color: "FFFFFF", fontFace: "Inter", bold: true, valign: "middle" });
        }
        else if (tid === "template-4") {
            slide.addImage({ path: bgDefault, x: 0, y: 0, w: '100%', h: '100%' });
            commonTitle(slide, f.title);
            slide.addText(f.paragraph1, { x: pX(88), y: pY(161), w: pX(845), h: pY(500), fontSize: fZ(22), color: "FFFFFF", fontFace: "Inter", valign: "top" });
            addFooter(slide);
        }
        else if (tid === "template-7") {
            slide.addImage({ path: bgDefault, x: 0, y: 0, w: '100%', h: '100%' });
            commonTitle(slide, f.title);
            slide.addText((f.paragraph1 || "") + "\n\n" + (f.paragraph2 || ""), { x: pX(88), y: pY(200), w: pX(860), h: pY(600), fontSize: fZ(22), color: "FFFFFF", fontFace: "Inter", valign: "top" });
            slide.addImage({ path: placeholder1, x: pX(973), y: 0, w: pX(947), h: pY(1080) });
            addFooter(slide);
        }
        else if (tid === "template-8") {
            slide.addImage({ path: bgDefault, x: 0, y: 0, w: '100%', h: '100%' });
            slide.addText(f.title, { x: pX(973 + 88), y: pY(61), w: pX(800), h: pY(60), fontSize: fZ(48), color: "FFFFFF", fontFace: "Inter", valign: "top" });
            slide.addText((f.paragraph1 || "") + "\n\n" + (f.paragraph2 || ""), { x: pX(1061), y: pY(200), w: pX(770), h: pY(600), fontSize: fZ(22), color: "FFFFFF", fontFace: "Inter", valign: "top" });
            slide.addImage({ path: placeholder2, x: 0, y: 0, w: pX(973), h: pY(1080) });
            addFooter(slide);
        }
        else if (tid === "template-6") {
            slide.addImage({ path: bgDefault, x: 0, y: 0, w: '100%', h: '100%' });
            commonTitle(slide, f.title);
            const bullets = [];
            for (let i = 1; i <= 13; i++) {
                if (f[`list_item_${i}`]) bullets.push(f[`list_item_${i}`]);
            }
            // First column
            if (bullets.length > 0) {
                const b1 = bullets.slice(0, 6).map(b => ({ text: b, options: { bullet: { type: 'number' } } }));
                slide.addText(b1, { x: pX(88), y: pY(161), w: pX(860), h: pY(600), fontSize: fZ(22), color: "FFFFFF", fontFace: "Inter", valign: "top" });
            }
            if (bullets.length > 6) {
                const b2 = bullets.slice(6, 13).map((b, i) => ({ text: b, options: { bullet: { type: 'number', numberStartAt: i + 7 } } }));
                slide.addText(b2, { x: pX(973), y: pY(161), w: pX(860), h: pY(600), fontSize: fZ(22), color: "FFFFFF", fontFace: "Inter", valign: "top" });
            }
            addFooter(slide);
        }
        else if (tid === "template-9") {
            slide.addImage({ path: bgDefault, x: 0, y: 0, w: '100%', h: '100%' });
            commonTitle(slide, f.title);
            const bullets = [];
            for (let i = 1; i <= 8; i++) {
                if (f[`bullet_${i}`]) bullets.push({ text: f[`bullet_${i}`], options: { bullet: true } });
            }
            if (bullets.length > 0) {
                slide.addText(bullets, { x: pX(88), y: pY(200), w: pX(1744), h: pY(600), fontSize: fZ(22), color: "FFFFFF", fontFace: "Inter", valign: "top", bullet: true });
            }
            addFooter(slide);
        }
        else if (tid === "template-10") {
            slide.addImage({ path: bgDefault, x: 0, y: 0, w: '100%', h: '100%' });
            commonTitle(slide, f.title);
            
            // Timeline line
            slide.addShape(pres.ShapeType.rect, { x: 0, y: pY(424), w: '100%', h: pY(3), fill: '4C4C4C' });
            
            for (let i = 1; i <= 4; i++) {
                const colsLeft = [210, 615, 1026, 1437];
                const cx = colsLeft[i-1];
                
                // Point
                slide.addShape(pres.ShapeType.oval, { x: pX(cx + 60), y: pY(377), w: pX(94), h: pY(94), line: { color: "333333", width: 1 }, fill: "1A1A1A" });
                slide.addShape(pres.ShapeType.oval, { x: pX(cx + 95), y: pY(412), w: pX(23), h: pY(23), fill: "FFE781" });
                
                slide.addText(f[`title_${i}`] || "", { x: pX(cx), y: pY(534), w: pX(298), h: pY(40), fontSize: fZ(28), color: "FFE781", fontFace: "Inter", bold: true });
                slide.addText(f[`text_${i}`] || "", { x: pX(cx), y: pY(580), w: pX(298), h: pY(200), fontSize: fZ(22), color: "D1D1D1", fontFace: "Inter", valign: "top" });
            }
            addFooter(slide);
        }
        else if (tid === "template-11") {
            slide.addImage({ path: bgDefault, x: 0, y: 0, w: '100%', h: '100%' });
            commonTitle(slide, f.title);
            
            const rows = [];
            // Header
            rows.push([
                { text: "", options: { fill: "0F1115", color: "D1D1D1", border: [null, null, {type:'solid', color:'333333'}, null] } },
                ...[f.col_1, f.col_2, f.col_3, f.col_4, f.col_5, f.col_6].map(c => ({ text: c || "", options: { fill: "0F1115", color: "D1D1D1", border: [null, null, {type:'solid', color:'333333'}, null] } }))
            ]);
            
            for (let i = 1; i <= 9; i++) {
                if (f[`row${i}_name`]) {
                    const rColor = i === 5 ? "FFE781" : "FFFFFF";
                    rows.push([
                        { text: f[`row${i}_name`], options: { color: rColor, align: 'left', border: [null, null, {type:'solid', color:'1A1A1A'}, null] } },
                        ...[f[`row${i}_c1`], f[`row${i}_c2`], f[`row${i}_c3`], f[`row${i}_c4`], f[`row${i}_c5`], f[`row${i}_c6`]].map(c => ({ text: c || "", options: { color: rColor, align: 'center', border: [null, null, {type:'solid', color:'1A1A1A'}, null] } }))
                    ]);
                }
            }
            
            slide.addTable(rows, { x: pX(88), y: pY(160), w: pX(1744), colW: [pX(1744)*0.3, pX(1744)*0.14, pX(1744)*0.14, pX(1744)*0.14, pX(1744)*0.14, pX(1744)*0.14], fontSize: fZ(24), fontFace: 'Inter' });
            
            addFooter(slide);
        }
    }
    
    // Trigger generic download inside browser
    return pres.writeFile({ fileName: downloadName });
}
