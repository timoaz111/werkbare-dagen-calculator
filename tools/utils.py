"""Gedeelde hulpfuncties voor Werkbare Dagen tools."""


def ms_to_beaufort(ms: float) -> int:
    """Converteert windsnelheid in m/s naar Beaufort schaal."""
    schaal = [0.3, 1.6, 3.4, 5.5, 8.0, 10.8, 13.9, 17.2, 20.8, 24.5, 28.5, 32.7]
    for bft, grens in enumerate(schaal):
        if ms < grens:
            return bft
    return 12
