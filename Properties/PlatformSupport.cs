// Projektweit markieren, dass die Assembly Windows (ab Win 6.1) benötigt.
// Entfernt CA1416-Warnungen für Windows-spezifische WinForms-/Forms-APIs.
using System.Runtime.Versioning;

[assembly: SupportedOSPlatform("windows6.1")]