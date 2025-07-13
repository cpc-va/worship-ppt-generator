
import React, { useState } from "react";
import PptxGenJS from "pptxgenjs";
import { Converter } from "opencc-js";
import { Dialog, DialogTitle, DialogContent, DialogActions, TextField, Box, Button, FormControl, InputLabel, MenuItem, Select, Typography, Grid } from "@mui/material";

const converter = Converter({ from: "tw", to: "cn" });

// TODO: support more background images
const backgrounds = [
  { name: "Background 1", url: "/bg/sun.jpg" },
  { name: "Background 2", url: "/bg/dark-purple.jpg" },
];

const chtFontFace = "Microsoft JhengHei"
// const chsFontFace = "Microsoft YaHei"
const chsFontFace = "Microsoft JhengHei"
const enFontFace = "Arial";
// TODO: make lyrics font size configurable
const coverFontSizeCh = 60;
const coverFontSizeEn = 36;
const lyricsFontSizeCh = 40;
const lyricsFontSizeEn = 28;
const footerFontSize = 12;
const blankLineHeight = 10;

function App() {
  const [songTitleCh, setSongTitleCh] = useState("");
  const [songTitleEn, setSongTitleEn] = useState("");
  const [credits, setCredits] = useState("");
  const [lyricsCh, setLyricsCh] = useState("");
  const [lyricsEn, setLyricsEn] = useState("");
  const [bgImage, setBgImage] = useState(backgrounds[0].url);
  const [fontColor, setFontColor] = useState("");
  const [chFontFace, setChFontFace] = useState("");
  const [errorOpen, setErrorOpen] = useState(false); // for modal

  const generatePPT = async (lang) => {
    const pptx = new PptxGenJS();
    const isSimplified = lang === "simp";
    setChFontFace(isSimplified ? chsFontFace : chtFontFace);

    // Add cover slide
    const coverSlide = pptx.addSlide();
    coverSlide.background = { path: window.location.origin + bgImage };
    coverSlide.addText(
      [
        {
          text: isSimplified ? converter(songTitleCh.trim()) : songTitleCh.trim(),
          options: {
            fontSize: coverFontSizeCh,
            fontFace: chFontFace,
            breakLine: true
          }
        },
        {
          text: songTitleEn,
          options: {
            fontSize: coverFontSizeEn,
            fontFace: enFontFace
          }
        }
      ],
      { x: 0.25, y: 0.4, w: "95%", h: 4.75, align: "center", color: fontColor, bold: true}
    );

    // prepare to add lyrics slides
    const blocksCh = lyricsCh.trim().split(/\n\s*\n/);
    const blocksEn = lyricsEn.trim().split(/\n\s*\n/); // blocks separated by double newlines
    let blockIndex = 0;
    for (let blockIndex = 0; blockIndex < blocksCh.length; blockIndex++) {
      // add new slide and set background
      const slide = pptx.addSlide();
      slide.background = { path: window.location.origin + bgImage };
      // check each block to make sure the # of lines between Chinese and English lyrics matches
      const chLines = blocksCh[blockIndex].trim().split(/\r?\n/).map((l) => l.trim());
      const enLines = blocksEn[blockIndex].trim().split(/\r?\n/).map((l) => l.trim());
      // Validate line count
      if (chLines.length !== enLines.length) {
        setErrorOpen(true);
        return;
      }
      // Build lyrics block
      const textBlocks = [];
      for (let i = 0; i < chLines.length; i++) {
        const ch = isSimplified ? converter(chLines[i]) : chLines[i];
        const en = enLines[i];
        // add chinese lyrics
        textBlocks.push(
          {
            text: ch,
            options: {
              align: "center",
              fontSize: lyricsFontSizeCh,
              color: fontColor,
              bold: true,
              fontFace: chFontFace,
              breakLine: true
            }
          }
        );
        // add english lyrics
        textBlocks.push(
          {
            text: en,
            options: {
              align: "center",
              fontSize: lyricsFontSizeEn,
              color: fontColor,
              fontFace: enFontFace,
              breakLine: true
            }
          }
        );
        // add blank line afterwards
        textBlocks.push(
          {
            text: " ",
            options: {
              align: "center",
              fontSize: blankLineHeight,
              fontFace: enFontFace,
              breakLine: true
            }
          }
        );
      }

      // Add lyrics text
      slide.addText(textBlocks, { x: 0.25, y: 0.4, w: "95%", h: 4.75, align: "center"});

      // Add footer (song name, credits)
      slide.addText(
        [
          {
            text: isSimplified ? converter(songTitleCh.trim()) : songTitleCh.trim(),
            options: {
              fontFace: chFontFace,
            }
          },
          {
            text: " " + songTitleEn,
            options: {
              fontFace: enFontFace
            }
          }
        ],
        { x: 0.2, y: 5.25, w: "48%", h: 0.25, align: "left", fontSize: footerFontSize, color: fontColor}
      );
      slide.addText(
        [
          {
            text: isSimplified ? converter(credits.trim()) : credits.trim(),
            options: {
              fontFace: chFontFace,
            }
          }
        ],
        { x: 5, y: 5.25, w: "48%", h: 0.25, align: "right", fontSize: footerFontSize, color: fontColor}
      );
    }
    const suffix = isSimplified ? "簡" : "繁";
    await pptx.writeFile(`${songTitleCh} ${songTitleEn} (${suffix}).pptx`);
  };

  return (
    <div>
      <Dialog open={errorOpen} onClose={() => setErrorOpen(false)}>
        <DialogTitle>Line Count Mismatch</DialogTitle>
        <DialogContent>
          <Typography>
            The number of Chinese and English lyrics lines must be the same.
          </Typography>
        </DialogContent>
        <DialogActions>
          <Button onClick={() => setErrorOpen(false)}>OK</Button>
        </DialogActions>
      </Dialog>
      <Grid container spacing={1}>
        <Grid item size={12} align="center">
          <Typography variant="h3" component="h1" gutterBottom>
            Worship Song PPT Generator
          </Typography>
        </Grid>
        <Grid item size={12}>
          <Typography variant="h6" fontStyle={"italic"}>
            注意：請用繁體中字輸入所有中文。<br/>
            NOTE: For all Chinese input, please use Traditionl Chinese.
          </Typography>
        </Grid>
        <Grid item size={6}>
          <TextField
            fullWidth
            label="歌名 (中文)"
            variant="outlined"
            margin="normal"
            placeholder="i.e. 祢真偉大"
            value={songTitleCh}
            onChange={(e) => setSongTitleCh(e.target.value)}
          />
        </Grid>
        <Grid item size={6}>
          <TextField
            fullWidth
            label="Song Title (English)"
            variant="outlined"
            margin="normal"
            helperText="Leave blank if there is no English title"
            placeholder="i.e. How Great Thou Art"
            value={songTitleEn}
            onChange={(e) => setSongTitleEn(e.target.value)}
          />
        </Grid>
        <Grid item size={12}>
          <TextField
            fullWidth
            label="版權 / Credits"
            variant="outlined"
            margin="normal"
            helperText="Leave blank for public domain songs or unknown sources"
            placeholder="i.e. 'By Hillsong', or '約書亞樂團版權所有'"
            value={credits}
            onChange={(e) => setCredits(e.target.value)}
          />
        </Grid>
        <Grid item size={12}>
          <Typography variant="body1" fontStyle={"italic"}>
            請於下面輸入歌詞。<br/>
            NOTE: When both Chinese and English lyrics are provided, the number of lines provided for Chinese lyrics must exactly match the number of lines provided for English lyrics.<br/>
            This is needed to make sure the English lyrics can be properly inserted under each Chinese lyrics line.
          </Typography>
        </Grid>
        <Grid item size={6}>
          <TextField
            fullWidth
            label="中文歌詞 (用空白行來分開不同投影片)"
            multiline
            rows={15}
            variant="outlined"
            margin="normal"
            helperText="用空白行來分開不同投影片"
            placeholder={`正歌第1句\n正歌第2句\n正歌第3句\n正歌第4句\n\nPre-Chorus第1句\nPre-Chorus第2句\n\n副歌第1句\n副歌第2句\n副歌第3句\n副歌第4句\n\nBridge第1句\nBridge第2句`}
            value={lyricsCh}
            onChange={(e) => setLyricsCh(e.target.value)}
          />
        </Grid>
        <Grid item size={6}>
          <TextField
            fullWidth
            label="Lyrics (use double newlines to separate slides)"
            multiline
            rows={15}
            variant="outlined"
            margin="normal"
            helperText="Use double newlines to separate slides. Leave blank if there are no English lyrics."
            placeholder={`Verse line 1\nVerse line 2\nVerse line 3\nVerse line 4\n\nPre-Chorus line 1\nPre-Chorus line 2\n\nChorus line 1\nChorus line 2\nChorus line 3\nChorus line 4\n\nBridge line 1\nBridge line 2`}
            value={lyricsEn}
            onChange={(e) => setLyricsEn(e.target.value)}
          />
        </Grid>
        <Grid item size={12}>
          <FormControl fullWidth margin="normal">
            <InputLabel>Background</InputLabel>
            <Select
              value={bgImage}
              label="Background"
              onChange={(e) => {setBgImage(e.target.value);setFontColor(e.target.value.includes("dark") ? "FFFFFF" : "000000");}}
            >
              {backgrounds.map((bg) => (
                <MenuItem key={bg.name} value={bg.url}>
                  <Box display="flex" alignItems="center" gap={1}>
                    <img
                      src={bg.url}
                      alt={bg.name}
                      width={80}
                      height={50}
                      style={{ objectFit: "cover", borderRadius: 4 }}
                    />
                    <Typography variant="body2">{bg.name}</Typography>
                  </Box>
                </MenuItem>
              ))}
            </Select>
          </FormControl>
        </Grid>
        <Grid item size={12} align="center">
          <Button
            variant="contained"
            color="primary"
            onClick={() => generatePPT("trad")}
            sx={{ mr: 2 }}
          >
            Download Traditional PPT
          </Button>

          <Button
            variant="contained"
            color="secondary"
            onClick={() => generatePPT("simp")}
          >
            Download Simplified PPT
          </Button>
        </Grid>
        <Grid item size={12}>
          <Typography variant="body2" align="right">
            Developed by Wah for CPC<br/>
            Last updated: 2025-07-12<br/>
            v1.0
          </Typography>
        </Grid>
      </Grid>
    </div>
  );
}

export default App;
