// Projektweit markieren, dass die Assembly Windows (ab Win 6.1) ben�tigt.
// Entfernt CA1416-Warnungen f�r Windows-spezifische WinForms-/Forms-APIs.
using System.Runtime.Versioning;

[assembly: SupportedOSPlatform("windows6.1")]