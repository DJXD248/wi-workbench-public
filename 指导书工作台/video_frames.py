# -*- coding: utf-8 -*-
"""视频抽帧模块：使用 imageio-ffmpeg 自带的 ffmpeg，无需用户安装任何东西。

用法：
    from video_frames import extract_frames
    frames, mode = extract_frames("操作视频.mp4", "素材图片/工序1", interval=3)
"""
from __future__ import annotations
import re
import subprocess
from pathlib import Path


def get_ffmpeg() -> str:
    import imageio_ffmpeg
    return imageio_ffmpeg.get_ffmpeg_exe()


def video_duration(video) -> float:
    ff = get_ffmpeg()
    p = subprocess.run([ff, "-hide_banner", "-i", str(video)],
                       capture_output=True, text=True, errors="replace")
    m = re.search(r"Duration:\s*(\d+):(\d+):(\d+(?:\.\d+)?)", p.stderr or "")
    if not m:
        return 0.0
    h, mi, s = int(m.group(1)), int(m.group(2)), float(m.group(3))
    return h * 3600 + mi * 60 + s


def _run_extract(ff, video, outdir, vf, max_frames, pattern):
    subprocess.run([ff, "-hide_banner", "-loglevel", "error", "-y",
                    "-i", str(video), "-vf", vf, "-fps_mode", "vfr",
                    "-frames:v", str(max_frames), "-q:v", "3", pattern],
                   capture_output=True, text=True, errors="replace")


def extract_frames(video, outdir, interval=3.0, scene_thresh=0.35,
                   max_frames=24, width=1280):
    """优先按画面突变挑帧；挑不足 3 帧就退回按固定间隔抽。

    返回 (帧文件列表, 模式说明文字)
    """
    video = Path(video)
    if not video.exists():
        raise FileNotFoundError(f"找不到视频: {video}")
    ff = get_ffmpeg()
    outdir = Path(outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    for old in outdir.glob("frame_*.jpg"):
        old.unlink()
    pattern = str(outdir / "frame_%03d.jpg")

    _run_extract(ff, video, outdir,
                 f"select='gt(scene,{scene_thresh})',scale={width}:-2",
                 max_frames, pattern)
    frames = sorted(outdir.glob("frame_*.jpg"))
    mode = "按画面切换挑帧"
    if len(frames) < 3:
        for f in frames:
            f.unlink()
        _run_extract(ff, video, outdir,
                     f"fps=1/{interval},scale={width}:-2",
                     max_frames, pattern)
        frames = sorted(outdir.glob("frame_*.jpg"))
        mode = f"每 {interval:g} 秒均匀截帧"
    if not frames:
        raise RuntimeError("抽帧失败：请确认视频文件完好、格式为常见 mp4/avi/mov")
    return frames, mode
