
import React, { useState } from "react";
import PptxGenJS from "pptxgenjs";
import { Converter } from "opencc-js";
import { Dialog, DialogTitle, DialogContent, DialogActions, TextField, Box, Button, FormControl, InputLabel, MenuItem, Select, Typography, Grid, Link, ListItem, List } from "@mui/material";

const converter = Converter({ from: "tw", to: "cn" });

// TODO: support more background images
const URI = "/worship-ppt-generator"
const backgrounds = [
  { name: "White (No Background)", url: URI + "/bg/white.jpg" },
  { name: "Black", url: URI + "/bg/dark-black.jpg" },
  { name: "Watercolor 1", url: URI + "/bg/watercolor-01.jpg" },
  { name: "Watercolor 2", url: URI + "/bg/watercolor-02.jpg" },
  { name: "Watercolor 3", url: URI + "/bg/watercolor-03.jpg" },
  { name: "Watercolor 4", url: URI + "/bg/watercolor-04.jpg" },
  { name: "Watercolor 5", url: URI + "/bg/watercolor-05.jpg" },
  { name: "Watercolor 6", url: URI + "/bg/watercolor-06.jpg" },
  { name: "Watercolor 7", url: URI + "/bg/watercolor-07.jpg" },
  { name: "Watercolor 8", url: URI + "/bg/watercolor-08.jpg" },
  { name: "Watercolor 9", url: URI + "/bg/watercolor-09.jpg" },
  { name: "Watercolor 10", url: URI + "/bg/watercolor-10.jpg" },
  { name: "Bubbles 1", url: URI + "/bg/bubbles-01.jpg" },
  { name: "Bubbles 2", url: URI + "/bg/bubbles-02.jpg" },
  { name: "Cloud 1", url: URI + "/bg/cloud-01.jpg" },
  { name: "Yellow 1", url: URI + "/bg/yellow-01.jpg" },
  { name: "Green 1", url: URI + "/bg/green-01.jpg" },
  { name: "Green 2", url: URI + "/bg/green-02.jpg" },
  { name: "Green 3", url: URI + "/bg/green-03.jpg" },
  { name: "Dark 1", url: URI + "/bg/dark-01.jpg" },
  { name: "Dark 2", url: URI + "/bg/dark-02.jpg" },
  { name: "Dark 3", url: URI + "/bg/dark-03.jpg" },
  { name: "Dark 4", url: URI + "/bg/dark-04.jpg" },
  { name: "Dark 5", url: URI + "/bg/dark-05.jpg" },
  { name: "Dark 6", url: URI + "/bg/dark-06.jpg" },
  { name: "Dark 7", url: URI + "/bg/dark-07.jpg" },
  { name: "Dark 8", url: URI + "/bg/dark-08.jpg" },
];

const chtFontFace = "Microsoft JhengHei"
// const chsFontFace = "Microsoft YaHei"
const chsFontFace = "Microsoft JhengHei"
const enFontFace = "Arial";
// TODO: make lyrics font size configurable
const coverFontSizeCh = 60;
const coverFontSizeEn = 36;
const lyricsFontSizeCh = 40;
const lyricsFontSizeEn = 24;
const footerFontSize = 12;
const blankLineHeight = 10;

function cleanChineseLine(line) {
  // Define Chinese & ASCII punctuation
  // const punctuation = /[，。！？：；「」『』（）—…、《》【】〈〉·!?,.:;"'()\[\]{}\-]/g;
  const punctuation = /[，。：；（）—…、《》【】〈〉·,.:;"'()\[\]{}\-]/g;

  return line.replace(punctuation, (match, offset, str) => {
    const isEnd = offset === str.length - 1;
    return isEnd ? "" : "   "; // remove at end, replace with 3 spaces elsewhere
  }).replace(/\s+/g, "   ");
}

function cleanEnglishLine(line) {
  return line.replace(/[.,?;:]+$/, "");
}

function App() {
  const [songTitleCh, setSongTitleCh] = useState("");
  const [songTitleEn, setSongTitleEn] = useState("");
  const [credits, setCredits] = useState("");
  const [lyricsCh, setLyricsCh] = useState("");
  const [lyricsEn, setLyricsEn] = useState("");
  const [bgImage, setBgImage] = useState(backgrounds[0].url);
  const [fontColor, setFontColor] = useState("");
  const [chFontFace, setChFontFace] = useState("");
  const [errorOpen, setErrorOpen] = useState(false);
  const [errorMsg, setErrorMsg] = useState("");
  const [infoOpen, setInfoOpen] = useState(false);

  const generatePPT = async (lang) => {

    var hasSongTitleCh = songTitleCh.trim()
    var hasSongTitleEn = songTitleEn.trim()
    var hasCredits = credits.trim()
    var hasLyricsCh = lyricsCh.trim()
    var hasLyricsEn = lyricsEn.trim()

    // basic input validation
    if (!hasSongTitleCh && !hasSongTitleEn) {
      setErrorMsg("Please enter at least either the Chinese or English song title.");
      setErrorOpen(true);
      return;
    }
    if (!hasLyricsCh) {
      setErrorMsg("Please enter Chinese lyrics before generating the slides.");
      setErrorOpen(true);
      return;
    }

    const pptx = new PptxGenJS();
    const isSimplified = lang === "simp";
    setChFontFace(isSimplified ? chsFontFace : chtFontFace);

    // Add cover slide
    const coverSlide = pptx.addSlide();
    coverSlide.background = { path: window.location.origin + bgImage };
    coverSlide.color = fontColor;
    const coverTextBlocks = [];
    if (hasSongTitleCh) {
      coverTextBlocks.push(
        {
          text: isSimplified ? converter(songTitleCh.trim()) : songTitleCh.trim(),
          options: {
            fontSize: coverFontSizeCh,
            fontFace: chFontFace,
            breakLine: true
          }
        }
      );
    }
    if (hasSongTitleEn) {
      coverTextBlocks.push(
        {
          text: songTitleEn,
          options: {
            fontSize: hasSongTitleCh ? coverFontSizeEn : coverFontSizeCh,
            fontFace: enFontFace
          }
        }
      );
    }
    coverSlide.addText(coverTextBlocks, { x: 0.25, y: 0.4, w: "95%", h: 4.75, align: "center", bold: true});

    // prepare to add lyrics slides
    const blocksCh = lyricsCh.trim().split(/\n\s*\n/); // blocks separated by double newlines
    const blocksEn = hasLyricsEn ? lyricsEn.trim().split(/\n\s*\n/) : [];
    let blockIndex = 0;
    for (let blockIndex = 0; blockIndex < blocksCh.length; blockIndex++) {
      // add new slide and set background
      const slide = pptx.addSlide();
      slide.color = fontColor;
      slide.background = { path: window.location.origin + bgImage };
      // check each block to make sure the # of lines between Chinese and English lyrics matches
      const chLines = blocksCh[blockIndex].trim().split(/\r?\n/).map((l) => l.trim());
      const enLines = hasLyricsEn ? blocksEn[blockIndex].trim().split(/\r?\n/).map((l) => l.trim()) : [];
      // Validate line count
      if (hasLyricsEn && chLines.length !== enLines.length) {
        setErrorMsg("The number of Chinese and English lyric lines must be the same.");
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
            text: cleanChineseLine(ch),
            options: {
              align: "center",
              fontSize: lyricsFontSizeCh,
              bold: true,
              fontFace: chFontFace,
              breakLine: true
            }
          }
        );
        if (hasLyricsEn) {
          // add english lyrics
          textBlocks.push(
            {
              text: cleanEnglishLine(en),
              options: {
                align: "center",
                fontSize: lyricsFontSizeEn,
                fontFace: enFontFace,
                breakLine: true
              }
            }
          );
        }
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
      const footerTextBlocks = [];
      if (hasSongTitleCh) {
        footerTextBlocks.push(
          {
            text: isSimplified ? converter(songTitleCh.trim()) : songTitleCh.trim(),
            options: {
              fontFace: chFontFace,
            }
          }
        );
      }
      if (hasSongTitleEn) {
        footerTextBlocks.push(
          {
            text: hasSongTitleCh ? " " + songTitleEn : songTitleEn,
            options: {
              fontFace: enFontFace
            }
          }
        );
      }
      slide.addText(footerTextBlocks, { x: 0.2, y: 5.25, w: "48%", h: 0.25, align: "left", fontSize: footerFontSize});
      if (hasCredits) {
        slide.addText(
          [
            {
              text: isSimplified ? converter(credits.trim()) : credits.trim(),
              options: {
                fontFace: chFontFace,
              }
            }
          ],
          { x: 5, y: 5.25, w: "48%", h: 0.25, align: "right", fontSize: footerFontSize}
        );
      }
    }
    const suffix = isSimplified ? "簡" : "繁";
    const songFileName = hasSongTitleCh ? songTitleCh + " " + songTitleEn : songTitleEn;
    await pptx.writeFile(`${songFileName} (${suffix}).pptx`);
  };

  return (
    <div>
      <Dialog open={errorOpen} onClose={() => setErrorOpen(false)}>
        <DialogTitle>Validation Error</DialogTitle>
        <DialogContent>
          <Typography>{errorMsg}</Typography>
        </DialogContent>
        <DialogActions>
          <Button onClick={() => setErrorOpen(false)} autoFocus>
            OK
          </Button>
        </DialogActions>
      </Dialog>
      <Dialog open={infoOpen} onClose={() => setInfoOpen(false)}>
        <DialogTitle>Why Use This?</DialogTitle>
        <DialogContent dividers>
          <Box display="flex" flexDirection="column" gap={2}>
            <Typography>
              This tool is created for CPC worship leaders to quickly prepare standardized PowerPoint slides by simply providing the following:
              <ul>
                <li>Song Title</li>
                <li>Lyrics Text</li>
                <li>Credits (Optional)</li>
              </ul>
              At its current state, this tool supports the following features:
              <ul>
                <li>Prepare bilingual lyrics slide (Chinese/English lyrics by sentence)</li>
                <li>Apply standard formatting (font face/size, spacing, stripping out extra punctuations, etc.)</li>
                <li>Support different backgrounds (font color will invert to white for dark backgrounds)</li>
                <li>Generate both Traditional and Simplified Chinese slides (source input should be in Traditional Chinese)</li>
              </ul>
            </Typography>
            <Typography fontWeight={600}>
              NOTE: This tool is designed to generate PowerPoint slides file for one song at a time; once you've generate all your individual slides files, you should make the corresponding updates and assemble them into a single PowerPoint file.
            </Typography>
          </Box>
        </DialogContent>
        <DialogActions>
          <Button onClick={() => setInfoOpen(false)}>Close</Button>
        </DialogActions>
      </Dialog>
      <Box width="100%">
        <Grid container spacing={1}>
          <Grid item size={{ xs: 12, sm: 12 }} align="center">
            <Typography variant="h3" component="h1" gutterBottom>
              Worship Song PPT Generator
            </Typography>
            <Link
              component="button"
              variant="h6"
              onClick={() => setInfoOpen(true)}
              sx={{ mt: 1, ml: 2 }}
            >
              Why use this?
            </Link>
          </Grid>
          <Grid item size={{ xs: 12, sm: 12 }}>
            <Typography variant="h6" fontStyle={"italic"}>
              注意：請用繁體中字輸入所有中文。<br/>
              NOTE: For all Chinese input, please use Traditionl Chinese.
            </Typography>
          </Grid>
          <Grid item size={{ xs: 12, sm: 6 }}>
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
          <Grid item size={{ xs: 12, sm: 6 }}>
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
          <Grid item size={{ xs: 12, sm: 12 }}>
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
          <Grid item size={{ xs: 12, sm: 12 }}>
            <Typography variant="body1" fontStyle={"italic"}>
              請於下面輸入歌詞。<br/>
              NOTE: When both Chinese and English lyrics are provided, the number of lines provided for Chinese lyrics must exactly match the number of lines provided for English lyrics.<br/>
              This is needed to make sure the English lyrics can be properly inserted under each Chinese lyrics line.
            </Typography>
          </Grid>
          <Grid item size={{ xs: 12, sm: 6 }}>
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
          <Grid item size={{ xs: 12, sm: 6 }}>
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
          <Grid item size={{ xs: 12, sm: 12 }}>
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
          <Grid item size={{ xs: 12, sm: 12 }}align="center">
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
          <Grid item size={{ xs: 12, sm: 12 }}>
            <Typography variant="body2" align="right">
              Developed by Wah for CPC<br/>
              Last updated: 2025-07-12<br/>
              v1.0
            </Typography>
          </Grid>
        </Grid>
      </Box>
    </div>
  );
}

export default App;
