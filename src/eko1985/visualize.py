"""Utilities for visualising management replays."""
from __future__ import annotations

from argparse import ArgumentParser
from collections import defaultdict
from math import nan
from pathlib import Path
from typing import Iterable, Sequence

import matplotlib.pyplot as plt

from .replay import excel_to_json, run_management_from_json

ENG_TO_SWE = {
    'Pine': 'Tall',
    'Spruce': 'Gran',
    'Birch': 'Björk',
    'Beech': 'Bok',
    'Oak': 'Ek',
    'Broadleaf': 'Öv.löv',
}

DEFAULT_METRICS: tuple[str, ...] = ('delta:BA', 'delta:N', 'delta:QMD', 'delta:VOL')

_PACKAGE_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_WORKBOOKS: tuple[Path, ...] = (
    _PACKAGE_ROOT / 'assets' / 'Output3.xlsx',
    _PACKAGE_ROOT / 'assets' / 'output4.xlsx',
    _PACKAGE_ROOT / 'assets' / 'Output2.xlsx',
)


def _align_species(snapshot: dict) -> dict[str, dict]:
    """Convert a snapshot's species keys from English to Swedish."""

    aligned: dict[str, dict] = {}
    for eng, values in (snapshot.get('species') or {}).items():
        swe = ENG_TO_SWE.get(eng, eng)
        aligned[swe] = values
    return aligned


def _parse_metric_spec(metric: str) -> tuple[str, str]:
    """Split metric specifications into their source and key components."""

    for delimiter in (':', '.'):
        if delimiter in metric:
            source, key = metric.split(delimiter, 1)
            return source.strip(), key.strip()
    return 'model', metric.strip()


def plot_replay_metrics(
    xls_paths: Iterable[str | Path],
    output_dir: str | Path,
    metrics: Sequence[str] = DEFAULT_METRICS,
) -> dict[str, list[Path]]:
    """Plot replay metrics for the given Excel workbooks.

    Parameters
    ----------
    xls_paths:
        Iterable of paths to Excel workbooks that define replay sequences.
    output_dir:
        Directory where generated figures should be written.
    metrics:
        Sequence of metric specifications to visualise. A metric can be provided
        either as a bare key (e.g. ``"BA"``) which will default to the
        ``"model"`` values, or as ``"source:key"`` / ``"source.key"`` to select
        an explicit source such as ``"expected:BA"``.

    Returns
    -------
    dict[str, list[Path]]
        Mapping of workbook stems to the saved figure paths.
    """

    out_path = Path(output_dir)
    out_path.mkdir(parents=True, exist_ok=True)

    saved: dict[str, list[Path]] = {}

    for raw_path in xls_paths:
        workbook_path = Path(raw_path)
        if not workbook_path.exists():
            continue

        replay_json = excel_to_json(str(workbook_path))
        snapshots = run_management_from_json(replay_json)
        if not snapshots:
            continue

        aligned_snaps = [_align_species(snapshot) for snapshot in snapshots]
        species_names: set[str] = set()
        for snap in aligned_snaps:
            species_names.update(snap.keys())

        if not species_names:
            continue

        species_order = sorted(species_names)
        event_labels = [
            snap.get('event') or f'Event {idx + 1}'
            for idx, snap in enumerate(snapshots)
        ]
        x_positions = list(range(len(event_labels)))

        workbook_saved: list[Path] = []

        for metric in metrics:
            metric_source, metric_key = _parse_metric_spec(metric)
            series_by_species: dict[str, list[float]] = defaultdict(list)
            for swe_name in species_order:
                for snap in aligned_snaps:
                    species_entry = snap.get(swe_name)
                    if not isinstance(species_entry, dict):
                        value = None
                    else:
                        source_block = species_entry.get(metric_source)
                        if not isinstance(source_block, dict):
                            value = None
                        else:
                            value = source_block.get(metric_key)
                    series_by_species[swe_name].append(
                        float(value) if value is not None else nan
                    )

            fig, ax = plt.subplots(figsize=(10, 5))
            for swe_name, series in series_by_species.items():
                ax.plot(
                    x_positions,
                    series,
                    marker='o',
                    label=swe_name,
                )

            ax.set_xticks(x_positions)
            ax.set_xticklabels(event_labels, rotation=45, ha='right')
            ax.set_xlabel('Händelse')
            label = f'{metric_source}:{metric_key}' if metric_source else metric_key
            if metric_source == 'delta':
                axis_label = f'{metric_key} (model - expected)'
                title = f'{workbook_path.stem} – {metric_key} (model - expected)'
            else:
                axis_label = metric_key
                title = f'{workbook_path.stem} – {label}'
            ax.set_ylabel(axis_label)
            ax.set_title(title)
            ax.grid(True, linestyle='--', linewidth=0.5, alpha=0.5)
            ax.legend()
            fig.tight_layout()

            safe_metric = metric.replace(':', '_')
            figure_path = out_path / f'{workbook_path.stem}_{safe_metric}.png'
            fig.savefig(figure_path)
            plt.close(fig)

            workbook_saved.append(figure_path)

        saved[workbook_path.stem] = workbook_saved

    return saved


def main(argv: Sequence[str] | None = None) -> int:
    """CLI entry point for running tests and plotting replay metrics."""

    parser = ArgumentParser(description='Run pytest and plot replay metrics.')
    parser.add_argument(
        '--output-dir',
        type=Path,
        default=_PACKAGE_ROOT / 'plots',
        help='Directory to store generated figures.',
    )
    parser.add_argument(
        '--workbooks',
        type=Path,
        nargs='*',
        default=list(DEFAULT_WORKBOOKS),
        help='Excel workbooks to replay (defaults to packaged assets).',
    )
    args = parser.parse_args(argv)

    import subprocess
    import sys

    result = subprocess.run([sys.executable, '-m', 'pytest', '-q'], check=False)
    if result.returncode != 0:
        return result.returncode

    plot_replay_metrics(args.workbooks, args.output_dir)
    print(f'Plots written to {args.output_dir}')
    return 0


__all__ = ['plot_replay_metrics', 'main']


if __name__ == "__main__":
    raise SystemExit(main())
