import sys
import click
from tqdm import tqdm

from core import ReportGenerator

from biz.radio import (
    RadioCheckGenerator
)

@click.command()
@click.option('-t', '--template-path')
@click.option('-f', '--file-path')
@click.option('-s', '--sheet')
@click.option('-d', '--target-dir')
@click.option('-m', '--mode')
def run(template_path: str, file_path: str, sheet: str, target_dir: str, mode: str):
    report_generator = None

    pbar = None

    def on_finish_one(number, total_cnt, success, exception):
        nonlocal pbar

        if pbar is None:
            pbar = tqdm(total=total_cnt)

        pbar.update(1)

    if mode == 'radio':
        report_generator = RadioCheckGenerator(template_path=template_path, on_finish_one=on_finish_one)

    if report_generator is None:
        sys.exit(f'Unknown mode:{mode}')

    report_generator.execute(file_path=file_path, target_dir=target_dir, sheet=sheet, global_data={})


if __name__ == '__main__':
    run()