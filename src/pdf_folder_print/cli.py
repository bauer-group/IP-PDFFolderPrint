"""
CLI mode (Rich console output).
"""

from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

from rich import box
from rich.console import Console
from rich.panel import Panel
from rich.progress import BarColumn, Progress, SpinnerColumn, TaskProgressColumn, TextColumn, TimeRemainingColumn
from rich.table import Table
from rich.tree import Tree

from . import __version__
from .core import collect_pdfs, get_pdf_info
from .printer import get_available_printers, get_default_printer, print_pdf_file


def cli_main() -> None:
    parser = argparse.ArgumentParser(
        prog="pdf_folder_print",
        description="PDF Folder Print – Structured PDF batch printing",
    )
    parser.add_argument("folder", nargs="?", help="Path to the folder containing PDFs")
    parser.add_argument("--gui", action="store_true", help="Launch GUI mode")
    parser.add_argument("--dry-run", action="store_true", help="List files without printing")
    parser.add_argument("--printer", default=None, help="Printer name (default: system default)")
    parser.add_argument("--dpi", type=int, default=None, help="Render DPI override (default: printer-native)")
    parser.add_argument("--delay", type=float, default=1.0, help="Delay between print jobs in seconds")
    parser.add_argument("--list-printers", action="store_true", help="List available printers and exit")
    parser.add_argument(
        "--yes", "-y", action="store_true", help="Start printing without confirmation prompt (for automation)"
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    args = parser.parse_args()

    console = Console()

    if args.gui:
        from .gui import gui_main

        gui_main()
        return

    if args.list_printers:
        default = get_default_printer()
        table = Table(title="Available Printers", box=box.ROUNDED, title_style="bold cyan")
        table.add_column("Printer", style="white")
        table.add_column("Status", style="dim")
        for name in get_available_printers():
            if name == default:
                table.add_row(f"[bold]{name}[/bold]", "[green]● default[/green]")
            else:
                table.add_row(name, "")
        console.print()
        console.print(table)
        console.print()
        sys.exit(0)

    if not args.folder:
        parser.print_help()
        sys.exit(1)

    root = Path(args.folder)
    if not root.is_dir():
        console.print(f"[bold red]ERROR:[/] Folder not found: {root}")
        sys.exit(1)

    if args.printer:
        available = get_available_printers()
        if args.printer not in available:
            console.print(f"[bold red]ERROR:[/] Printer [bold]'{args.printer}'[/] not found.\n")
            for p in available:
                console.print(f"  • {p}")
            sys.exit(1)
        printer_name = args.printer
    else:
        printer_name = get_default_printer()

    pdfs = collect_pdfs(root)
    if not pdfs:
        console.print("[yellow]No PDF files found.[/]")
        sys.exit(0)

    groups: dict[str, list[Path]] = {}
    for group, pdf_path in pdfs:
        groups.setdefault(group, []).append(pdf_path)

    dpi_label = f"{args.dpi} (manual)" if args.dpi else "printer-native"
    mode_label = "[yellow]DRY-RUN[/yellow] (preview only)" if args.dry_run else "[green bold]PRINT[/green bold]"

    info_table = Table(box=None, show_header=False, padding=(0, 2))
    info_table.add_column("Key", style="dim", no_wrap=True)
    info_table.add_column("Value", style="bold")
    info_table.add_row("Folder", root.name)
    info_table.add_row("Printer", printer_name)
    info_table.add_row("Render DPI", dpi_label)
    info_table.add_row("Scaling", "Fit-to-Page + auto orientation")
    info_table.add_row("Mode", mode_label)
    info_table.add_row("Files", f"{len(pdfs)} in {len(groups)} groups")

    console.print()
    console.print(Panel(info_table, title="[bold]PDF Folder Print[/]", border_style="cyan", expand=False))

    tree = Tree(f"[bold cyan]{root.name}[/]")
    for gname, gpaths in groups.items():
        if gname == "(Root)":
            for p in gpaths:
                tree.add(f"[dim]{p.name}[/]")
        else:
            branch = tree.add(f"[yellow]{gname}/[/] [dim]({len(gpaths)} files)[/]")
            for p in gpaths:
                branch.add(f"[dim]{p.name}[/]")

    console.print()
    console.print(tree)
    console.print()

    if not args.dry_run and not args.yes:
        answer = console.input(f"[bold]Start printing {len(pdfs)} files? [y/N][/] ").strip().lower()
        if answer not in ("y", "yes", "j", "ja"):
            console.print("[dim]Cancelled.[/]")
            sys.exit(0)
        console.print()

    ok_count = 0
    fail_count = 0
    total_pages = 0
    current_group = None
    unreadable: list[str] = []

    with Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        BarColumn(bar_width=30),
        TaskProgressColumn(),
        TextColumn("•"),
        TimeRemainingColumn(),
        console=console,
        transient=False,
    ) as progress:

        task = progress.add_task("[cyan]Processing...", total=len(pdfs))

        for i, (group, pdf_path) in enumerate(pdfs):
            if group != current_group:
                current_group = group
                progress.console.print(f"\n[bold yellow]📁 {group}[/]")

            if args.dry_run:
                n, w, h = get_pdf_info(pdf_path)
                if n > 0:
                    orient = "landscape" if w > h else "portrait"
                    orient_color = "magenta" if w > h else "blue"
                    progress.console.print(
                        f"   [dim]📄[/] {pdf_path.name}  "
                        f"[dim]→[/]  {n} page{'s' if n != 1 else ''}  "
                        f"[dim]{w:.0f}×{h:.0f}mm[/]  "
                        f"[{orient_color}]{orient}[/]"
                    )
                    ok_count += 1
                else:
                    progress.console.print(f"   [dim]📄[/] {pdf_path.name}  [red]→ unreadable[/]")
                    fail_count += 1
                    unreadable.append(pdf_path.name)
            else:
                progress.update(task, description=f"[cyan]{pdf_path.name}")
                ok, pages, err = print_pdf_file(pdf_path, printer_name, args.dpi)
                if ok:
                    progress.console.print(
                        f"   [green]✓[/] {pdf_path.name}  "
                        f"[dim]{pages} page{'s' if pages != 1 else ''}[/]"
                    )
                    ok_count += 1
                    total_pages += pages
                else:
                    progress.console.print(f"   [bold red]✗[/] {pdf_path.name}  [red]{err}[/]")
                    fail_count += 1
                if i < len(pdfs) - 1 and args.delay > 0:
                    time.sleep(args.delay)

            progress.update(task, advance=1)

        progress.update(task, description="[green]Done")

    console.print()

    if args.dry_run:
        style = "yellow" if fail_count == 0 else "red"
        result_table = Table(box=box.ROUNDED, border_style=style, title="[bold]Dry-Run Summary[/]")
        result_table.add_column("", style="dim")
        result_table.add_column("", style="bold")
        result_table.add_row("Files", f"{ok_count}")
        if unreadable:
            result_table.add_row("Unreadable", f"[bold red]{fail_count} files[/]")
        result_table.add_row("Groups", f"{len(groups)}")
        result_table.add_row(
            "Status",
            "[yellow]Preview only – nothing printed[/]"
            if fail_count == 0
            else "[red]✗ Unreadable PDFs detected[/]",
        )
        console.print(result_table)
        if unreadable:
            console.print()
            for name in unreadable:
                console.print(f"  [bold red]✗[/] [red]{name}[/]")
    else:
        style = "green" if fail_count == 0 else "red"
        result_table = Table(box=box.ROUNDED, border_style=style, title="[bold]Print Summary[/]")
        result_table.add_column("", style="dim")
        result_table.add_column("", style="bold")
        result_table.add_row("Printed", f"[green]{ok_count} files ({total_pages} pages)[/]")
        if fail_count:
            result_table.add_row("Failed", f"[bold red]{fail_count} files[/]")
        result_table.add_row("Printer", printer_name)
        result_table.add_row("Status", "[green]✓ Complete[/]" if fail_count == 0 else "[red]✗ Errors occurred[/]")
        console.print(result_table)

    console.print()

    # "Press Enter" is only useful when the window would close immediately,
    # i.e. for the packaged EXE with an attached console. Regular terminal
    # sessions and automation (Task Scheduler, pipes) must not block.
    if not args.dry_run and getattr(sys, "frozen", False) and sys.stdin is not None and sys.stdin.isatty():
        console.input("[dim]Press Enter to exit...[/]")

    sys.exit(1 if fail_count > 0 else 0)
