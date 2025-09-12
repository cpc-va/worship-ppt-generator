import { useState } from "react";
import PptxGenJS from "pptxgenjs";
import { Converter } from "opencc-js";
import { Dialog, DialogTitle, DialogContent, DialogActions, TextField, Box, Button, FormControl, InputLabel, MenuItem, Select, Typography, Grid, Link, FormLabel, RadioGroup, FormControlLabel, Radio, Tooltip, IconButton, Alert } from "@mui/material";
import { backgrounds, chtFontFace, chsFontFace, enFontFace, footerFontSize, blankLineHeight, lyricsFontSizeEnSec, lyricsFontSizeEnPri, lyricsFontSizeChSec, lyricsFontSizeChPri, coverFontSizeEnPri, coverFontSizeChPri, coverFontSizeEnSec, coverFontSizeChSec, lyricsFontSizeChPriSmaller, lyricsFontSizeChPriSmallest, lyricsFontSizeEnPriSmaller, lyricsFontSizeEnPriSmallest, lyricsFontSizeEnSecSmallest, lyricsFontSizeEnSecSmaller, lyricsFontSizeChSecSmallest } from "../constants";
import InfoOutlinedIcon from "@mui/icons-material/InfoOutlined";

const t2sConverter = Converter({ from: "tw", to: "cn" });
const s2tConverter = Converter({ from: "cn", to: "tw" });

function cleanChineseLine(line) {
  // Define Chinese & ASCII punctuation
  // const punctuation = /[，。！？：；「」『』（）—…、《》【】〈〉·!?,.:;"'()\[\]{}\-]/g;
  const punctuation = /[，。：；（）—…、《》【】〈〉·,.:;"'()\[\]{}\-]/g;

  return line !== undefined ?
    line.replace(punctuation, (offset, str) => {
      const isEnd = offset === str.length - 1;
      return isEnd ? "" : "   "; // remove at end, replace with 3 spaces elsewhere
    }).replace(/\s+/g, "   ")
    :
    "";
}

function cleanEnglishLine(line) {
  return line !== undefined ? line.replace(/[.,?;:]+$/, "") : "";
}

function WorshipPptGen() {
    
  const [primaryLang, setPrimaryLang] = useState("ch");
  const [songTitlePri, setSongTitlePri] = useState("");
  const [songTitleSec, setSongTitleSec] = useState("");
  const [credits, setCredits] = useState("");
  const [lyricsPri, setLyricsPri] = useState("");
  const [lyricsSec, setLyricsSec] = useState("");
  const [bgImage, setBgImage] = useState(backgrounds[0].url);
  const [fontColor, setFontColor] = useState("");
  const [chFontFace, setChFontFace] = useState("");
  const [errorOpen, setErrorOpen] = useState(false);
  const [errorMsg, setErrorMsg] = useState("");
  const [infoOpen, setInfoOpen] = useState(false);
  const [whatsNew, setWhatsNew] = useState(false);
  const [chFontSizeScale, setChFontSizeScale] = useState("standard");
  const [enFontSizeScale, setEnFontSizeScale] = useState("standard");
  const [simpChInputWarning, setSimpChInputWarning] = useState(false);

  // pass in 1 for primary, otherwise secondary
  function isPrmLangCh(primSec) {
    if (primSec == 1) return primaryLang === 'ch';
    return primaryLang !== 'ch';
  }
  function isSimplified(downloadLang) {
    return downloadLang === 'simp';
  }
  function getLang(primSec) {
    return isPrmLangCh(primSec) ? 'Chinese' : 'English';
  }
  function getSongTitleText(downloadLang, lineNo) {
    return primaryLang === 'ch' ?
      (lineNo == 1 ? getConvertedLine(downloadLang, songTitlePri.trim()) : songTitleSec.trim())
      :
      (lineNo == 1 ? songTitlePri.trim() : getConvertedLine(downloadLang, songTitleSec.trim()));
  }
  function getCoverFontSize(lineNo) {
    return primaryLang === 'ch' ?
      (lineNo == 1 ? coverFontSizeChPri : coverFontSizeEnSec)
      :
      (lineNo == 1 ? coverFontSizeEnPri : coverFontSizeChSec);
  }
  function getLyricsFontSize(lineNo) {
    return primaryLang === 'ch' ?
      (lineNo == 1 ? getLyricsFontSizeChPri() : getLyricsFontSizeEnSec())
      :
      (lineNo == 1 ? getLyricsFontSizeEnPri() : getLyricsFontSizeChSec());
  }
  function getLyricsFontSizeChPri() {
    if (chFontSizeScale === 'smaller') return lyricsFontSizeChPriSmaller;
    else if (chFontSizeScale === 'smallest') return lyricsFontSizeChPriSmallest;
    else return lyricsFontSizeChPri;
  }
  function getLyricsFontSizeChSec() {
    if (chFontSizeScale === 'smaller') return lyricsFontSizeChSecSmaller;
    else if (chFontSizeScale === 'smallest') return lyricsFontSizeChSecSmallest;
    else return lyricsFontSizeChSec;
  }
  function getLyricsFontSizeEnPri() {
    if (enFontSizeScale === 'smaller') return lyricsFontSizeEnPriSmaller;
    else if (enFontSizeScale === 'smallest') return lyricsFontSizeEnPriSmallest;
    else return lyricsFontSizeEnPri;
  }
  function getLyricsFontSizeEnSec() {
    if (enFontSizeScale === 'smaller') return lyricsFontSizeEnSecSmaller;
    else if (enFontSizeScale === 'smallest') return lyricsFontSizeEnSecSmallest;
    else return lyricsFontSizeEnSec;
  }
  function getFontFace(primSec) {
    return isPrmLangCh(primSec) ? 'chFontFace' : 'enFontFace';
  }
  function getSongLabel(primSec) {
    return isPrmLangCh(primSec) ? '歌名 (中文)' : 'Song Title (English)';
  }
  function getSongLabelSecHelpText() {
    return "Leave blank if there is no " + getLang(2) + " title";
  }
  function getPlaceHolder(primSec) {
    return isPrmLangCh(primSec) === 'ch' ? 'i.e. 祢真偉大' : 'i.e. How Great Thou Art';
  }
  function getLyricsLabel(primSec) {
    return isPrmLangCh(primSec) ? '中文歌詞 (用空白行來分開不同投影片)' : 'Lyrics (use double newlines to separate slides)';
  }
  function getLyricsHelperText(primSec) {
    var helperText = isPrmLangCh(primSec) ? '用空白行來分開不同投影片。' : 'Use double newlines to separate slides.';
    if (primSec == 2) {
      helperText += primaryLang === 'ch' ? ' Leave blank if there are no English lyrics.' : ' 如沒有中文歌詞，請留空。';
    }
    return helperText;
  }
  function getLyricsPlaceholder(primSec) {
    return isPrmLangCh(primSec) ?
      '正歌第1句\n正歌第2句\n正歌第3句\n正歌第4句\n\nPre-Chorus第1句\nPre-Chorus第2句\n\n副歌第1句\n副歌第2句\n副歌第3句\n副歌第4句\n\nBridge第1句\nBridge第2句'
      :
      'Verse line 1\nVerse line 2\nVerse line 3\nVerse line 4\n\nPre-Chorus line 1\nPre-Chorus line 2\n\nChorus line 1\nChorus line 2\nChorus line 3\nChorus line 4\n\nBridge line 1\nBridge line 2';
  }
  function getConvertedLine(downloadLang, line) {
    return isSimplified(downloadLang) ? t2sConverter(line) : (detectSimplifiedChineseInput(line) ? s2tConverter(line) : line);
  }
  function getLyricsLine(downloadLang, primSec, lyricsLine) {
    return isPrmLangCh(primSec) ?
      cleanChineseLine(getConvertedLine(downloadLang, lyricsLine))
      :
      cleanEnglishLine(lyricsLine);
  }
  function detectSimplifiedChineseInput(input) {
    if (input === undefined) return true;
    // Detect Simplified Chinese by comparing with converted version
    var converted = s2tConverter(input);
     return input !== converted;
  }

  const generatePPT = async (lang) => {

    var hasSongTitlePri = songTitlePri.trim()
    var hasSongTitleSec = songTitleSec.trim()
    var hasCredits = credits.trim()
    var hasLyricsPri = lyricsPri.trim()
    var hasLyricsSec = lyricsSec.trim()
    
    var songTitleCh = isPrmLangCh(1) ? songTitlePri : songTitleSec;
    var lyricsCh = isPrmLangCh(1) ? lyricsPri : lyricsSec;
    if (detectSimplifiedChineseInput(songTitleCh) || detectSimplifiedChineseInput(lyricsCh) || detectSimplifiedChineseInput(credits)) {
      setSimpChInputWarning(true);
    } else {
      setSimpChInputWarning(false);
    }

    // basic input validation
    if (!hasSongTitlePri && !hasSongTitleSec) {
      setErrorMsg("Please enter at least either the Chinese or English song title.");
      setErrorOpen(true);
      return;
    }
    if (!hasLyricsPri) {
      setErrorMsg("Please enter at least the " + getLang(1) + " lyrics before generating the slides.");
      setErrorOpen(true);
      return;
    }

    const pptx = new PptxGenJS();
    setChFontFace(isSimplified(lang) ? chsFontFace : chtFontFace);

    // Add cover slide
    const coverSlide = pptx.addSlide();
    coverSlide.background = { path: window.location.origin + bgImage };
    coverSlide.color = fontColor;
    const coverTextBlocks = [];
    if (hasSongTitlePri) {
      coverTextBlocks.push(
        {
          text: getSongTitleText(lang, 1),
          options: {
            fontSize: getCoverFontSize(1),
            fontFace: getFontFace(1),
            breakLine: true
          }
        }
      );
    }
    if (hasSongTitleSec) {
      coverTextBlocks.push(
        {
          text: getSongTitleText(lang, 2),
          options: {
            fontSize: getCoverFontSize(2),
            fontFace: getFontFace(2)
          }
        }
      );
    }
    coverSlide.addText(coverTextBlocks, { x: 0.25, y: 0.4, w: "95%", h: 4.75, align: "center", bold: true});

    // prepare to add lyrics slides
    const blocksPri = lyricsPri.trim().split(/\n\s*\n/); // blocks separated by double newlines
    const blocksSec = hasLyricsSec ? lyricsSec.trim().split(/\n\s*\n/) : [];
    let blockIndex = 0;
    for (let blockIndex = 0; blockIndex < blocksPri.length; blockIndex++) {
      // add new slide and set background
      const slide = pptx.addSlide();
      slide.color = fontColor;
      slide.background = { path: window.location.origin + bgImage };
      // check each block to make sure the # of lines between Chinese and English lyrics matches
      const priLines = blocksPri[blockIndex].trim().split(/\r?\n/).map((l) => l.trim());
      const secLines = hasLyricsSec ? blocksSec[blockIndex].trim().split(/\r?\n/).map((l) => l.trim()) : [];
      // Validate line count
      if (hasLyricsSec && priLines.length !== secLines.length) {
        setErrorMsg("The number of " + getLang(1) + " and " + getLang(2) + " lyric lines must be the same.");
        setErrorOpen(true);
        return;
      }
      // Build lyrics block
      const textBlocks = [];
      for (let i = 0; i < priLines.length; i++) {
        const priLyricsLine = getLyricsLine(lang, 1, priLines[i] ?? '');
        const secLyricsLine = getLyricsLine(lang, 2, secLines[i] ?? '');
        // add primary lyrics
        textBlocks.push(
          {
            text: priLyricsLine,
            options: {
              align: "center",
              fontSize: getLyricsFontSize(1),
              bold: true,
              fontFace: getFontFace(1),
              breakLine: true
            }
          }
        );
        if (hasLyricsSec) {
          // add secondary lyrics
          textBlocks.push(
            {
              text: secLyricsLine,
              options: {
                align: "center",
                fontSize: getLyricsFontSize(2),
                fontFace: getFontFace(2),
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
      if (hasSongTitlePri) {
        footerTextBlocks.push(
          {
            text: getConvertedLine(lang, songTitlePri.trim()),
            options: {
              fontFace: getFontFace(1),
            }
          }
        );
      }
      if (hasSongTitleSec) {
        footerTextBlocks.push(
          {
            text: hasSongTitlePri ? " " + getConvertedLine(lang, songTitleSec.trim()) : getConvertedLine(lang, songTitleSec.trim()),
            options: {
              fontFace: getFontFace(2)
            }
          }
        );
      }
      slide.addText(footerTextBlocks, { x: 0.2, y: 5.25, w: "48%", h: 0.25, align: "left", fontSize: footerFontSize});
      if (hasCredits) {
        slide.addText(
          [
            {
              text: getConvertedLine(lang, credits.trim()),
              options: {
                fontFace: chFontFace,
              }
            }
          ],
          { x: 5, y: 5.25, w: "48%", h: 0.25, align: "right", fontSize: footerFontSize}
        );
      }
    }
    const suffix = isSimplified(lang) ? "簡" : "繁";
    const songFileName = hasSongTitlePri ? songTitlePri + " " + songTitleSec : songTitleSec;
    await pptx.writeFile(`${songFileName} (${suffix}).pptx`);
  };

  function swapContents() {
    // swap song title
    var existingSongTitlePri = songTitlePri;
    var existingSongTitleSec = songTitleSec;
    setSongTitlePri(existingSongTitleSec);
    setSongTitleSec(existingSongTitlePri);
    // swap lyrics
    var existingLyricsPri = lyricsPri;
    var existingLyricsSec = lyricsSec;
    setLyricsPri(existingLyricsSec);
    setLyricsSec(existingLyricsPri);
  }

  const handleLangChange = (event) => {
    setPrimaryLang(event.target.value);
    // swap any existing contents that have been entered
    swapContents();
  };

  const handleChFontSizeScaleChange = (event) => {
    setChFontSizeScale(event.target.value);
  };

  const handleEnFontSizeScaleChange = (event) => {
    setEnFontSizeScale(event.target.value);
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
            <Typography>
              NOTE:
              <ul>
                <li>This tool is designed to generate PowerPoint slides file for <b>one song at a time;</b> once you've generate all your individual slides files, you should make the corresponding updates and assemble them into a single PowerPoint file.</li>
                <li>This slide generator is currently only intended for generating lyrics slides, but <b>not for Bible scripture slidses</b>. Please use your existing scripture slide templates to prepare them, if you need to insert them into your worship set.</li>
              </ul>
            </Typography>
          </Box>
        </DialogContent>
        <DialogActions>
          <Button onClick={() => setInfoOpen(false)}>Close</Button>
        </DialogActions>
      </Dialog>
      <Dialog open={whatsNew} onClose={() => setWhatsNew(false)}>
        <DialogTitle>What's New?</DialogTitle>
        <DialogContent dividers>
          <Box display="flex" flexDirection="column" gap={2}>
            <Typography>
              v1.1 (9/12/2025)
              <ul>
                <li><b>Add support for primary language toggle between Chinese and English:</b> you can now enter English lyrics as the primary language, where the English lyrics will first show up on the top of a lyrics line set.</li>
                <li><b>Support Simplified Chinese input:</b> the system will automatically detect Simplified Chinese input and try to convert it back to Traditional Chinese when generating such slides, although the conversion is not 100% since multiple Traditional Chinese characters can be mapped to a single Simplified Chinese character. As a result, it's still recommended to use Traditional Chinese for inputting all Chinese characters; the system will display a small warning banner if Simplified Chinese input is detected.</li>
                <li><b>Support adjustable lyrics font size:</b> both Chinese and English fonts can be independently adjusted to Standard, Smaller or Smallest font size to reduce line wrapping due to long sentences.</li>
                <li><b>Development updates:</b> performed some code cleanup and refactoring.</li>
              </ul>
              v1.0 (7/12/2025)
              <ul>
                <li><b>Generate bilingual lyrics slide</b> (Chinese/English lyrics by sentence)</li>
                <li><b>Apply standard formatting</b> (font face/size, spacing, stripping out extra punctuations, etc.)</li>
                <li><b>Support different backgrounds</b> (font color will invert to white for dark backgrounds)</li>
                <li><b>Generate both Traditional and Simplified Chinese slides</b> (source input should be in Traditional Chinese)</li>
              </ul>
            </Typography>
          </Box>
        </DialogContent>
        <DialogActions>
          <Button onClick={() => setWhatsNew(false)}>Close</Button>
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
              [Why use this?]
            </Link>
            <Link
              component="button"
              variant="h6"
              onClick={() => setWhatsNew(true)}
              sx={{ mt: 1, ml: 2 }}
            >
              [What's New?]
            </Link>
          </Grid>
          <Grid item size={{ xs: 12, sm: 12 }}>
            <Typography variant="h6" fontStyle={"italic"}>
              注意：為了避免字體轉換問題，建議使用繁體中字輸入所有中文。<br/>
              NOTE: It's recommended to use Traditional Chinese for all Chinese-character input to avoid potential conversion problems.
            </Typography>
          </Grid>
          <Grid size={{ xs: 12, sm: 12}}>
            <Box
              sx={{
                border: "1px solid #ccc",
                borderRadius: 2,
                p: 2,
                mt: 2
              }}
            >
              <FormControl component="fieldset">
                <Box sx={{ display: "flex", alignItems: "center", mb: 1 }}>
                  <FormLabel component="legend">主唱語言 Primary Language</FormLabel>
                  <Tooltip title="Determines which is the main language that appears on the top of each lyrics line group.">
                    <IconButton size="small" sx={{ p: 0.2 }}>
                      <InfoOutlinedIcon fontSize="small" />
                    </IconButton>
                  </Tooltip>
                </Box>
                <RadioGroup
                  row
                  name="primary-language"
                  value={primaryLang}
                  onChange={handleLangChange}
                >
                  <FormControlLabel
                      value="ch"
                      control={<Radio />}
                      label="Chinese"
                  />
                  <FormControlLabel
                      value="en"
                      control={<Radio />}
                      label="English"
                  />
                </RadioGroup>
              </FormControl>
            </Box>
          </Grid>
          <Grid item size={{ xs: 12, sm: 6 }}>
            <TextField
              fullWidth
              label={getSongLabel(1)}
              variant="outlined"
              margin="normal"
              placeholder={getPlaceHolder(1)}
              value={songTitlePri}
              onChange={(e) => setSongTitlePri(e.target.value)}
            />
          </Grid>
          <Grid item size={{ xs: 12, sm: 6 }}>
            <TextField
              fullWidth
              label={getSongLabel(2)}
              variant="outlined"
              margin="normal"
              helperText={getSongLabelSecHelpText()}
              placeholder={getPlaceHolder(2)}
              value={songTitleSec}
              onChange={(e) => setSongTitleSec(e.target.value)}
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
              label={getLyricsLabel(1)}
              multiline
              rows={15}
              variant="outlined"
              margin="normal"
              helperText={getLyricsHelperText(1)}
              placeholder={getLyricsPlaceholder(1)}
              value={lyricsPri}
              onChange={(e) => setLyricsPri(e.target.value)}
            />
          </Grid>
          <Grid item size={{ xs: 12, sm: 6 }}>
            <TextField
              fullWidth
              label={getLyricsLabel(2)}
              multiline
              rows={15}
              variant="outlined"
              margin="normal"
              helperText={getLyricsHelperText(2)}
              placeholder={getLyricsPlaceholder(2)}
              value={lyricsSec}
              onChange={(e) => setLyricsSec(e.target.value)}
            />
          </Grid>
          <Grid item size={{ xs: 12, sm: 6 }}>
            <FormControl fullWidth>
              <InputLabel id="font-size-label">{getLang(1)} Font Size</InputLabel>
              <Select
                labelId="font-size-label"
                value={primaryLang === 'ch' ? chFontSizeScale : enFontSizeScale}
                onChange={primaryLang === 'ch' ? handleChFontSizeScaleChange : handleEnFontSizeScaleChange}
              >
                <MenuItem value="standard">Standard</MenuItem>
                <MenuItem value="smaller">Smaller</MenuItem>
                <MenuItem value="smallest">Smallest</MenuItem>
              </Select>
            </FormControl>
          </Grid>
          <Grid item size={{ xs: 12, sm: 6 }}>
            <FormControl fullWidth>
              <InputLabel id="font-size-label">{getLang(2)} Font Size</InputLabel>
              <Select
                labelId="font-size-label"
                value={primaryLang !== 'ch' ? chFontSizeScale : enFontSizeScale}
                onChange={primaryLang !== 'ch' ? handleChFontSizeScaleChange : handleEnFontSizeScaleChange}
              >
                <MenuItem value="standard">Standard</MenuItem>
                <MenuItem value="smaller">Smaller</MenuItem>
                <MenuItem value="smallest">Smallest</MenuItem>
              </Select>
            </FormControl>
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
          {simpChInputWarning && (
            <Alert severity="warning" sx={{ mb: 2 }}>
              Your input contains Simplified Chinese characters.
              While the system will still try to perform characters conversion for you, 
              it is recommended to use Traditional Chinese input to avoid potential conversion problems.
            </Alert>
          )}
          <Grid item size={{ xs: 12, sm: 12 }}align="center">
            <Button
              variant="contained"
              color="primary"
              onClick={() => generatePPT("trad")}
              sx={{ mr: 2 }}
            >
              Download PPT (Traditional Chinese)
            </Button>

            <Button
              variant="contained"
              color="secondary"
              onClick={() => generatePPT("simp")}
            >
              Download PPT (Simplified Chinese)
            </Button>
          </Grid>
          <Grid item size={{ xs: 12, sm: 12 }}>
            <Typography variant="body2" align="right">
              Developed by Wah for CPC<br/>
              Last updated: 2025-09-12<br/>
              v1.1
            </Typography>
          </Grid>
        </Grid>
      </Box>
    </div>
  );
}

export default WorshipPptGen;