![image](https://github.com/user-attachments/assets/b3d5481f-5e75-4626-bd70-f636d41a54d5)
# EveryDiscord - A Discord Client for Legacy Windows

A simple, fast, and reliable native Discord client for legacy Windows versions, including Windows XP and 2000.

## Features

- Native VB6/tB implementation for optimal performance on older systems, as a native Win32-based software unlike the Electron Discord desktop client
- Built with VbAsyncSocket and VBWebSocket for stable connections
- Lightweight alternative to the modern Discord client
- Compatible with Windows XP and other legacy Windows versions
- Basic Discord functionality in a minimal package
- No need for SSE2 instructions
- Plugins built in like CatBox-Based 200mb file uploads, upcoming free emoji support(F@ken1tro Backport), theming etc.

## Compatibility

Tested on:
- Windows XP x86 (SP3), Modern AMD + Pentium 4
- Windows XP x64 (SP2)
- Windows 11
- Windows 7 x64 (SP1)
- *Windows 95(Failed, tester writing wrong token or incompability is unclear though)
- *Windows 98 + Pentium 2 as per the TLS library, VbAsyncSocket

## Technical Details

**Development Language:** Visual Basic 6.0  & twinBASIC
**Key Dependencies:**
- VbAsyncSocket (for network operations with TLS 1.3 without even needing SSE2)
- VBWebSocket (for WebSocket connections)
- VBA-FastJSON
- VBCCR 1.8

## Installation

1. Download the latest release from Releases Page
2. Launch EveryDiscord and log in with your Discord token, OAuth2 will be implemented later

## Building from Source

Requirements:
- Visual Basic 6.0 IDE or twinBASIC IDE
- Prayers
