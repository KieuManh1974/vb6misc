Attribute VB_Name = "basDXSound"
Option Explicit
Global m_dx As New DirectX7
Global m_ds As DirectSound
Public m_dsPrimaryBuffer As DirectSoundBuffer
Public m_dsListener As DirectSound3DListener
Public m_dsBuffer(2) As DirectSoundBuffer
Public m_ds3DBuffer(2) As DirectSound3DBuffer
Global m_dsCapture As DirectSoundCapture
Public WaveFMT As WAVEFORMATEX
Public wBuffer() As Byte
Public CurPos As DSCURSORS


