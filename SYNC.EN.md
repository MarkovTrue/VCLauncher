[Русский](SYNC.md) | [English](SYNC.EN.md)

# Offset detection

`Sync.exe` finds how far one release of a movie is shifted relative to the other. VCLauncher runs it on "Compare" in "Auto" mode and passes the result to Video-compare.

## By audio

1. The first 5 minutes are skipped, then one minute of audio is taken from both files.
2. Loudness envelopes are compared: they barely depend on the codec or channel count.
3. The best-matching shift is searched for. A weak or ambiguous peak is rejected.
4. If the first tracks don't match, other track pairs are tried, same language first.
5. A shift under half a frame is checked on two more segments. It is usually a track delay left by remuxing.

## By video

Fallback method. It runs when audio finds no match. It also verifies audio shifts over 10 s. If the results differ by more than 0.5 s, the offset is rejected.

1. Frames are downscaled and converted to grayscale. Brightness is normalized so color grading doesn't interfere.
2. An abrupt change between neighboring frames counts as a cut.
3. Cuts from both files are matched in pairs. The shift that aligns the most cuts is chosen.

## Settings

The `[Settings]` section of `VCLauncher.ini`:

| Key | Default | Purpose |
|---|---|---|
| `SyncSkipSec` | 300 | Skipped from the start, s |
| `SyncTimeoutSec` | 60 | Time limit per search, s |
| `SyncSuspectMs` | 10000 | Larger shifts are verified by scene changes, ms |
| `SyncVerifyTolMs` | 500 | Allowed mismatch when verifying, ms |

## Limitations

- The offset is measured near the start. If scenes are cut or added, it differs later on.
- Without a common audio track and with very different color grading, no offset is found. Set it manually then.
