from __future__ import annotations

import argparse
import base64
import html
import math
import os
import re
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable

import pandas as pd


SPECIAL_RULES = {
    "victor caua gomes maciel": {"exclude_from_general_rate": {"resolucao"}},
}


DISPLAY_LABELS = {
    "Relatorio": "Relatório",
    "Periodo": "Período",
    "Analise": "Análise",
    "Criterios": "Critérios",
    "criterios": "critérios",
    "criterio": "critério",
    "Conclusoes": "Conclusões",
    "Recomendacoes": "Recomendações",
    "Introducao": "Introdução",
    "Observacoes": "Observações",
    "emissao": "emissão",
    "Observacao Interna": "Observação Interna",
    "Resolucao": "Resolução",
    "Cadastro Zendesk": "Cadastro Zendesk",
    "Avaliacao Geral do Atendimento": "Avaliação Geral do Atendimento",
    "Criterio": "Critério",
    "Nao Cumprido (%)": "Não Cumprido (%)",
    "Cumprido (%)": "Cumprido (%)",
    "Total de Registros": "Total de Registros",
    "Colaborador": "Colaborador",
    "Taxa de Cumprimento Geral": "Taxa de Cumprimento Geral",
    "Obs. Interna (C/NC)": "Obs. Interna (C/NC)",
    "Resolucao (C/NC)": "Resolução (C/NC)",
    "Zendesk (C/NC)": "Zendesk (C/NC)",
    "Total de Atendimentos": "Total de Atendimentos",
    "Nao informado": "Não informado",
    "Nao cumprido": "Não cumprido",
}


@dataclass
class CriterionSummary:
    name: str
    met: int
    not_met: int
    total: int

    @property
    def met_pct(self) -> float:
        return (self.met / self.total * 100.0) if self.total else 0.0

    @property
    def not_met_pct(self) -> float:
        return (self.not_met / self.total * 100.0) if self.total else 0.0


def normalize_text(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return text


def slugify_header(value: str) -> str:
    text = normalize_text(value)
    text = re.sub(r"[^a-z0-9]+", "_", text).strip("_")
    return text


def find_column(columns: Iterable[str], *candidates: str) -> str | None:
    normalized = {slugify_header(col): col for col in columns}
    for candidate in candidates:
        match = normalized.get(slugify_header(candidate))
        if match:
            return match
    return None


def convert_yes_x(value: object) -> int | None:
    normalized = normalize_text(value)
    if not normalized:
        return None
    if normalized == "sim":
        return 1
    if normalized == "x":
        return 0
    return None


def categorize_general_evaluation(value: object) -> str | None:
    normalized = normalize_text(value)
    if not normalized:
        return None
    if "correto" in normalized or "acolhedor" in normalized:
        return "correto"
    if "atencao" in normalized or "atenção" in str(value).lower():
        return "atencao"
    return None


def format_count_pair(met: int, not_met: int) -> str:
    return f"{met}/{not_met}"


def format_ticket_value(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return "-"
    text = str(value).strip()
    if not text:
        return "-"
    if text.endswith(".0"):
        maybe_number = text[:-2]
        if maybe_number.isdigit():
            return maybe_number
    return text


def display_label(text: object) -> str:
    return DISPLAY_LABELS.get(str(text), str(text))


def display_free_text(text: str) -> str:
    rendered = text
    for source, target in DISPLAY_LABELS.items():
        rendered = rendered.replace(source, target)
    return rendered


def ensure_dirs(base_path: Path) -> tuple[Path, Path]:
    data_dir = base_path / "data"
    output_dir = base_path / "docs"
    mpl_config_dir = base_path / ".mplconfig"
    data_dir.mkdir(exist_ok=True)
    output_dir.mkdir(exist_ok=True)
    mpl_config_dir.mkdir(exist_ok=True)
    return data_dir, output_dir


def load_dataframe(csv_path: Path, encoding: str) -> pd.DataFrame:
    df = pd.read_csv(csv_path, encoding=encoding, sep=None, engine="python")
    df.columns = [str(col).strip() for col in df.columns]
    return df


def infer_period_from_dataframe(df: pd.DataFrame) -> str | None:
    date_column = find_column(df.columns, "data", "date")
    if not date_column:
        return None

    dates = pd.to_datetime(df[date_column], dayfirst=True, errors="coerce").dropna()
    if dates.empty:
        return None

    start = dates.min().strftime("%d/%m/%Y")
    end = dates.max().strftime("%d/%m/%Y")
    return f"{start} a {end}"


def prepare_dataframe(df: pd.DataFrame) -> tuple[pd.DataFrame, dict[str, str]]:
    column_map = {
        "colaborador": find_column(
            df.columns,
            "acolhedor",
            "colaborador",
            "nome do acolhedor",
            "nome",
        ),
        "observacao_interna": find_column(
            df.columns,
            "observacao interna",
            "observação interna",
        ),
        "resolucao": find_column(df.columns, "resolucao", "resolução"),
        "zendesk": find_column(
            df.columns,
            "cadastro zendesk",
            "zendesk",
            "cadastro no zendesk",
            "zendesk cadastro",
            "zendesk (cadastro)",
        ),
        "avaliacao_geral": find_column(
            df.columns,
            "avaliacao geral do atendimento",
            "avaliação geral do atendimento",
            "avaliacao geral",
        ),
        "observacoes": find_column(df.columns, "observacoes", "observações"),
        "tipo_atendimento": find_column(
            df.columns,
            "tipo de atendimento",
            "canal",
        ),
        "chat": find_column(df.columns, "chat"),
        "redes_sociais": find_column(df.columns, "redes sociais"),
        "email": find_column(df.columns, "email"),
        "ligacao": find_column(df.columns, "ligacao", "ligação"),
    }

    missing = [key for key, value in column_map.items() if key in {
        "colaborador",
        "observacao_interna",
        "resolucao",
        "zendesk",
        "avaliacao_geral",
    } and value is None]
    if missing:
        raise ValueError(
            "Colunas obrigatorias nao encontradas: " + ", ".join(sorted(missing))
        )

    prepared = df.copy()
    prepared["__colaborador__"] = prepared[column_map["colaborador"]].fillna("").astype(str)
    prepared["__obs_interna__"] = prepared[column_map["observacao_interna"]].apply(convert_yes_x)
    prepared["__resolucao__"] = prepared[column_map["resolucao"]].apply(convert_yes_x)
    prepared["__zendesk__"] = prepared[column_map["zendesk"]].apply(convert_yes_x)
    prepared["__avaliacao_categoria__"] = prepared[column_map["avaliacao_geral"]].apply(
        categorize_general_evaluation
    )
    prepared["__observacoes__"] = (
        prepared[column_map["observacoes"]].fillna("").astype(str)
        if column_map["observacoes"]
        else ""
    )
    if column_map["tipo_atendimento"]:
        prepared["__tipo_atendimento__"] = (
            prepared[column_map["tipo_atendimento"]].fillna("Não informado").astype(str)
        )
    else:
        def infer_channel(row: pd.Series) -> str:
            channel_columns = [
                ("Chat", column_map["chat"]),
                ("Redes Sociais", column_map["redes_sociais"]),
                ("Email", column_map["email"]),
                ("Liga\u00e7\u00e3o", column_map["ligacao"]),
            ]
            found = []
            for label, col in channel_columns:
                if not col:
                    continue
                value = row.get(col)
                if pd.notna(value) and str(value).strip() != "":
                    found.append(label)
            return ", ".join(found) if found else "Não informado"

        prepared["__tipo_atendimento__"] = prepared.apply(infer_channel, axis=1)
    return prepared, column_map


def summarize_binary_criterion(df: pd.DataFrame, column: str, name: str) -> CriterionSummary:
    valid = df[column].dropna()
    total = int(valid.count())
    met = int(valid.sum()) if total else 0
    not_met = total - met
    return CriterionSummary(name=name, met=met, not_met=not_met, total=total)


def summarize_general_evaluation(df: pd.DataFrame) -> CriterionSummary:
    valid = df["__avaliacao_categoria__"].dropna()
    total = int(valid.count())
    met = int((valid == "correto").sum())
    not_met = int((valid == "atencao").sum())
    return CriterionSummary(
        name="Avaliacao Geral do Atendimento",
        met=met,
        not_met=not_met,
        total=total,
    )


def build_collaborator_table(df: pd.DataFrame) -> pd.DataFrame:
    records: list[dict[str, object]] = []
    grouped = df.groupby("__colaborador__", dropna=False)
    for collaborator, group in grouped:
        if not str(collaborator).strip():
            continue

        normalized_name = normalize_text(collaborator)
        exclusions = SPECIAL_RULES.get(normalized_name, {}).get("exclude_from_general_rate", set())

        metrics = {
            "observacao_interna": summarize_binary_criterion(
                group, "__obs_interna__", "Observacao Interna"
            ),
            "resolucao": summarize_binary_criterion(group, "__resolucao__", "Resolucao"),
            "zendesk": summarize_binary_criterion(group, "__zendesk__", "Cadastro Zendesk"),
        }

        general_met = 0
        general_total = 0
        for key, summary in metrics.items():
            if key in exclusions:
                continue
            general_met += summary.met
            general_total += summary.total

        overall_rate = (general_met / general_total * 100.0) if general_total else math.nan

        evaluations = group["__avaliacao_categoria__"].dropna()
        evaluation_icons = []
        if (evaluations == "correto").any():
            evaluation_icons.append("OK")
        if (evaluations == "atencao").any():
            evaluation_icons.append("ATENCAO")

        records.append(
            {
                "Colaborador": collaborator,
                "Taxa de Cumprimento Geral": overall_rate,
                "Obs. Interna (C/NC)": format_count_pair(
                    metrics["observacao_interna"].met, metrics["observacao_interna"].not_met
                ),
                "Resolucao (C/NC)": format_count_pair(
                    metrics["resolucao"].met, metrics["resolucao"].not_met
                ),
                "Zendesk (C/NC)": format_count_pair(
                    metrics["zendesk"].met, metrics["zendesk"].not_met
                ),
                "Total de Atendimentos": int(len(group)),
            }
        )

    result = pd.DataFrame(records)
    if not result.empty:
        result = result.sort_values(
            by=["Taxa de Cumprimento Geral", "Colaborador"],
            ascending=[False, True],
            na_position="last",
        )
    return result


def build_observation_groups(df: pd.DataFrame) -> list[dict[str, object]]:
    groups: dict[str, dict[str, object]] = {}
    for _, row in df.iterrows():
        category = row["__avaliacao_categoria__"]
        collaborator = row["__colaborador__"]
        if collaborator not in groups:
            groups[collaborator] = {
                "colaborador": collaborator,
                "atencao": 0,
                "ok": 0,
                "itens": [],
            }
        if not collaborator.strip():
            continue

        channel_parts = []
        channel_value = str(row.get("__tipo_atendimento__", "")).strip()
        if channel_value and channel_value != "Nao informado":
            channel_parts.append(channel_value)

        ticket = ""
        for col in ("Chat", "Redes Sociais", "Email", "Ligação", "Ligacao"):
            if col in row.index and pd.notna(row[col]) and str(row[col]).strip() != "":
                ticket = format_ticket_value(row[col])
                break

        observation = str(row.get("__observacoes__", "")).strip()

        if category == "atencao":
            groups[collaborator]["atencao"] += 1
            emoji = "⚠️"
            status_label = "Aten\u00e7\u00e3o"
        else:
            groups[collaborator]["ok"] += 1
            emoji = "✅"
            status_label = "Correto"

        groups[collaborator]["itens"].append(
            {
                "ticket": ticket or "-",
                "categoria": category or "correto",
                "emoji": emoji,
                "status_label": status_label,
                "texto": observation if observation else "Sem observa\u00e7\u00e3o adicional.",
                "canal": ", ".join(channel_parts) if channel_parts else "Nao informado",
            }
        )

    for group in groups.values():
        group["itens"] = sorted(
            group["itens"],
            key=lambda item: (
                item["ticket"],
                item["canal"],
                item["texto"],
            ),
        )

    return sorted(
        groups.values(),
        key=lambda item: item["colaborador"],
    )


def build_channel_summary(df: pd.DataFrame) -> pd.DataFrame:
    counts = (
        df["__tipo_atendimento__"]
        .fillna("Não informado")
        .astype(str)
        .value_counts()
        .rename_axis("Tipo de Atendimento")
        .reset_index(name="Total")
    )
    return counts


def plot_consolidated_chart(summaries: list[CriterionSummary], output_path: Path) -> None:
    labels = [
        display_label(summary.name).replace("Avaliação Geral do Atendimento", "Avaliação Geral")
        for summary in summaries
    ]
    met_values = [summary.met_pct for summary in summaries]
    not_met_values = [summary.not_met_pct for summary in summaries]

    fig, ax = plt.subplots(figsize=(10.5, 5.8))
    fig.patch.set_facecolor("#fbf7f1")
    ax.set_facecolor("#fffdf9")
    positions = list(range(len(labels)))

    ax.barh(positions, not_met_values, color="#eadfd2", height=0.56, label="Não cumprido")
    ax.barh(positions, met_values, color="#2f7d4a", height=0.56, label="Cumprido")

    for idx, (met, not_met) in enumerate(zip(met_values, not_met_values)):
        ax.text(min(met - 2, 96), idx, f"{met:.1f}%", va="center", ha="right", color="white", fontsize=10, fontweight="bold")
        ax.text(101, idx, f"{not_met:.1f}%", va="center", ha="left", color="#8a4a30", fontsize=9)

    ax.set_xlim(0, 108)
    ax.set_yticks(positions)
    ax.set_yticklabels(labels)
    ax.set_xlabel("Percentual")
    ax.set_title("Cumprimento por critério", loc="left", fontsize=16, fontweight="bold")
    ax.grid(axis="x", linestyle="--", alpha=0.18)
    ax.spines[["top", "right", "left", "bottom"]].set_visible(False)
    ax.legend(loc="lower right", frameon=False)
    fig.tight_layout()
    fig.savefig(output_path, dpi=200)
    plt.close(fig)


def plot_collaborator_chart(df: pd.DataFrame, output_path: Path) -> None:
    if df.empty:
        return

    chart_df = df.copy()
    colors = []
    for value in chart_df["Taxa de Cumprimento Geral"]:
        if value >= 95:
            colors.append("#2f7d4a")
        elif value >= 80:
            colors.append("#d89b3c")
        else:
            colors.append("#b55239")

    fig, ax = plt.subplots(figsize=(11.5, max(6, len(chart_df) * 0.5)))
    fig.patch.set_facecolor("#fbf7f1")
    ax.set_facecolor("#fffdf9")
    ax.barh(chart_df["Colaborador"], chart_df["Taxa de Cumprimento Geral"], color=colors, height=0.62)
    for idx, value in enumerate(chart_df["Taxa de Cumprimento Geral"]):
        ax.text(value + 1.2, idx, f"{value:.1f}%", va="center", ha="left", fontsize=10, color="#3d352f")
    ax.set_xlim(0, 100)
    ax.set_xlabel("Taxa de Cumprimento Geral (%)")
    ax.set_title("Desempenho geral por colaborador", loc="left", fontsize=16, fontweight="bold")
    ax.grid(axis="x", linestyle="--", alpha=0.18)
    ax.spines[["top", "right", "left", "bottom"]].set_visible(False)
    ax.invert_yaxis()
    fig.tight_layout()
    fig.savefig(output_path, dpi=200)
    plt.close(fig)


def dataframe_to_markdown(df: pd.DataFrame, float_cols: set[str] | None = None) -> str:
    if df.empty:
        return "_Sem dados disponiveis._"

    float_cols = float_cols or set()
    rendered = df.copy()
    for col in float_cols:
        if col in rendered.columns:
            rendered[col] = rendered[col].apply(
                lambda value: "-" if pd.isna(value) else f"{float(value):.1f}%"
            )

    def escape_cell(value: object) -> str:
        text = str(value)
        text = text.replace("|", "\\|").replace("\n", " ").strip()
        return text

    headers = [escape_cell(col) for col in rendered.columns]
    rows = [[escape_cell(value) for value in row] for row in rendered.fillna("-").values.tolist()]
    separator = ["---"] * len(headers)
    markdown_lines = [
        "| " + " | ".join(headers) + " |",
        "| " + " | ".join(separator) + " |",
    ]
    markdown_lines.extend("| " + " | ".join(row) + " |" for row in rows)
    return "\n".join(markdown_lines)


def dataframe_to_html_table(
    df: pd.DataFrame,
    float_cols: set[str] | None = None,
    table_class: str = "data-table",
) -> str:
    if df.empty:
        return '<p class="empty-state">Sem dados disponiveis.</p>'

    float_cols = float_cols or set()
    rendered = df.copy()
    rendered = rendered.rename(columns=lambda col: display_label(col))
    rendered = rendered.map(lambda value: display_label(value) if isinstance(value, str) else value)
    for col in float_cols:
        if col in rendered.columns:
            rendered[col] = rendered[col].apply(
                lambda value: "-" if pd.isna(value) else f"{float(value):.1f}%"
            )

    headers = "".join(f"<th>{html.escape(str(col))}</th>" for col in rendered.columns)
    rows = []
    for values in rendered.fillna("-").values.tolist():
        cells = "".join(
            f"<td>{html.escape(str(value)).replace(chr(10), '<br>')}</td>" for value in values
        )
        rows.append(f"<tr>{cells}</tr>")

    return (
        f'<table class="{table_class}">'
        f"<thead><tr>{headers}</tr></thead>"
        f"<tbody>{''.join(rows)}</tbody>"
        "</table>"
    )


def image_to_data_uri(path: Path) -> str:
    mime = "image/png" if path.suffix.lower() == ".png" else "application/octet-stream"
    encoded = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:{mime};base64,{encoded}"


def build_kpi_cards(
    consolidated: list[CriterionSummary],
    collaborator_df: pd.DataFrame,
    channel_df: pd.DataFrame,
) -> list[dict[str, str]]:
    consolidated_map = {item.name: item for item in consolidated}
    total_attendances = int(channel_df["Total"].sum()) if not channel_df.empty else 0
    return [
        {
            "label": "Cumprimento Geral",
            "value": f'{consolidated_map["Avaliacao Geral do Atendimento"].met_pct:.1f}%',
            "note": "Percentual de atendimentos classificados como corretos",
        },
        {
            "label": "Observação Interna",
            "value": f'{consolidated_map["Observacao Interna"].met_pct:.1f}%',
            "note": "Cumprimento do registro interno",
        },
        {
            "label": "Total de Atendimentos",
            "value": str(total_attendances),
            "note": "Volume total analisado no período",
        },
    ]


def build_conclusions(consolidated: list[CriterionSummary]) -> list[str]:
    quantitative = [item for item in consolidated if item.name != "Avaliacao Geral do Atendimento"]
    ordered = sorted(quantitative, key=lambda item: item.met_pct, reverse=True)
    best = ordered[0]
    improvement_targets = ordered[-2:] if len(ordered) >= 2 else ordered
    improvement_targets = sorted(improvement_targets, key=lambda item: item.met_pct)
    improvement_text = " e ".join(item.name for item in improvement_targets)

    conclusion_lines = [
        "## 5. Conclusoes e Recomendacoes",
        "",
        f"- **Pontos Fortes:** O critério de {display_label(best.name)} continua apresentando um bom desempenho geral, indicando que a equipe tem sido eficaz nesse aspecto do atendimento.",
        f"- **Pontos de Melhoria:** {display_free_text(improvement_text)} são os pontos que necessitam de maior atenção, com percentuais de cumprimento que indicam oportunidades de melhoria.",
        "- **Recomendacoes:**",
        "1.1 Treinamento Contínuo: Reforçar treinamentos sobre a importância e o processo correto de registro na Zendesk e a qualidade das observações internas.",
        "",
    ]
    return conclusion_lines


def generate_markdown_report(
    period: str,
    title: str,
    consolidated: list[CriterionSummary],
    collaborator_df: pd.DataFrame,
    observation_groups: list[dict[str, object]],
    output_dir: Path,
) -> str:
    emission_date = datetime.now().strftime("%d/%m/%Y")
    consolidated_df = pd.DataFrame(
        [
            {
                "Criterio": item.name,
                "Cumprido (%)": f"{item.met_pct:.1f}%",
                "Nao Cumprido (%)": f"{item.not_met_pct:.1f}%",
                "Total de Registros": item.total,
            }
            for item in consolidated
        ]
    )

    collaborator_table_df = collaborator_df.sort_values(by="Colaborador", ascending=True).reset_index(drop=True)

    report_lines = [
        f"# {title} - {period}",
        "",
        f"**Data de Emissao:** {emission_date}",
        "",
        "## 1. Introducao",
        "",
        "Este relatorio apresenta o desempenho dos atendimentos online com foco nos criterios Observacao Interna, Resolucao, Cadastro Zendesk e Avaliacao Geral do Atendimento.",
        "",
        "## 2. Analise Consolidada dos Criterios",
        "",
        "A tabela abaixo resume a taxa de cumprimento por criterio no periodo analisado.",
        "",
        f"![Analise Consolidada de Criterios]({(output_dir / 'analise_criterios.png').name})",
        "",
        dataframe_to_markdown(consolidated_df),
        "",
        "## 3. Analise de Desempenho por Colaborador",
        "",
        "O ranking a seguir mostra a taxa geral de cumprimento por colaborador com base nos criterios quantitativos aplicaveis.",
        "",
        f"![Desempenho Geral por Colaborador]({(output_dir / 'desempenho_colaboradores.png').name})",
        "",
        dataframe_to_markdown(
            collaborator_table_df,
            float_cols={"Taxa de Cumprimento Geral"},
        ),
        "",
        "**Legenda:** `C/NC` = Cumprido / Nao Cumprido.",
        "",
        "## 4. Observacoes Adicionais",
        "",
        "### Observações Detalhadas",
        "",
    ]

    if observation_groups:
        for group in observation_groups:
            report_lines.append(
                f"**{group['colaborador']}** - {group['atencao']} ponto(s) de atencao, {group['ok']} ponto(s) positivo(s)"
            )
            for item in group["itens"]:
                report_lines.append(
                    f"- {item['emoji']} Ticket {item['ticket']} | {item['canal']} | {item['status_label']} | {item['texto']}"
                )
            report_lines.append("")
    else:
        report_lines.append("_Sem observacoes detalhadas no periodo._")

    report_lines.extend(
        [
            "",
            *build_conclusions(consolidated),
            "Atenciosamente, Hemily Marques",
            "",
        ]
    )
    return "\n".join(report_lines)


def generate_html_report(
    period: str,
    title: str,
    consolidated: list[CriterionSummary],
    collaborator_df: pd.DataFrame,
    channel_df: pd.DataFrame,
    observation_groups: list[dict[str, object]],
    output_dir: Path,
) -> str:
    emission_date = datetime.now().strftime("%d/%m/%Y")
    consolidated_df = pd.DataFrame(
        [
            {
                "Criterio": item.name,
                "Cumprido (%)": f"{item.met_pct:.1f}%",
                "Nao Cumprido (%)": f"{item.not_met_pct:.1f}%",
                "Total de Registros": item.total,
            }
            for item in consolidated
        ]
    )
    collaborator_table_df = collaborator_df.sort_values(by="Colaborador", ascending=True).reset_index(drop=True)
    kpis = build_kpi_cards(consolidated, collaborator_df, channel_df)
    consolidated_chart_uri = image_to_data_uri(output_dir / "analise_criterios.png")
    collaborator_chart_uri = image_to_data_uri(output_dir / "desempenho_colaboradores.png")
    observations_html = ""
    for group in observation_groups:
        status_class = "attention" if int(group["atencao"]) > 0 else "positive"
        status_label = (
            "Com pontos de aten\u00e7\u00e3o" if int(group["atencao"]) > 0 else "Sem pontos de aten\u00e7\u00e3o"
        )
        items_html = ""
        for item in group["itens"]:
            item_class = (
                "attention" if item["categoria"] == "atencao" else
                "positive" if item["categoria"] == "correto" else
                "info"
            )
            items_html += (
                f'<li class="obs-item {item_class}">'
                f'<div class="obs-meta">'
                f'<span class="obs-emoji">{html.escape(str(item["emoji"]))}</span>'
                f'<span class="obs-pill">{html.escape(str(item["status_label"]))}</span>'
                f'<span class="obs-ticket">Ticket {html.escape(str(item["ticket"]))}</span>'
                f'<span class="obs-channel">{html.escape(str(item["canal"]))}</span>'
                "</div>"
                f'<p>{html.escape(str(item["texto"]))}</p>'
                "</li>"
            )
        observations_html += (
            f'<article class="observation-card {status_class}">'
            f'<div class="obs-header">'
            f'<div><h3>{html.escape(str(group["colaborador"]))}</h3>'
            f'<p>{int(group["ok"])} ponto(s) positivo(s) e {int(group["atencao"])} ponto(s) de aten\u00e7\u00e3o</p></div>'
            f'<span class="obs-status">{status_label}</span>'
            "</div>"
            f'<ul class="observation-list">{items_html}</ul>'
            "</article>"
        )
    if not observations_html:
        observations_html = '<p class="empty-state">Sem observa\u00e7\u00f5es detalhadas no período.</p>'

    kpi_cards_html = "".join(
        (
            '<article class="kpi-card">'
            f'<p class="kpi-label">{html.escape(card["label"])}</p>'
            f'<p class="kpi-value">{html.escape(card["value"])}</p>'
            f'<p class="kpi-note">{html.escape(card["note"])}</p>'
            "</article>"
        )
        for card in kpis
    )

    quantitative = [item for item in consolidated if item.name != "Avaliacao Geral do Atendimento"]
    ordered_quantitative = sorted(quantitative, key=lambda item: item.met_pct, reverse=True)
    best_name = ordered_quantitative[0].name if ordered_quantitative else "-"
    improvement_targets = ordered_quantitative[-2:] if len(ordered_quantitative) >= 2 else ordered_quantitative
    improvement_targets = sorted(improvement_targets, key=lambda item: item.met_pct)
    improvement_text = " e ".join(item.name for item in improvement_targets) if improvement_targets else "-"

    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>{html.escape(display_free_text(title))} - {html.escape(period)}</title>
  <style>
    :root {{
      --bg: #f5f1e8;
      --panel: #fffdf8;
      --ink: #202020;
      --muted: #6f6a62;
      --accent: #c46a2f;
      --accent-dark: #83411a;
      --ok: #2f7d4a;
      --warn: #b55239;
      --line: #e6ddcf;
      --shadow: 0 18px 40px rgba(72, 45, 24, 0.08);
      --soft-green: #edf7f0;
      --soft-red: #fbefea;
      --soft-gold: #faf2df;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: Georgia, "Times New Roman", serif;
      color: var(--ink);
      background:
        radial-gradient(circle at top left, rgba(196,106,47,0.12), transparent 26%),
        linear-gradient(180deg, #f8f4ec 0%, var(--bg) 100%);
    }}
    .page {{
      max-width: 1200px;
      margin: 0 auto;
      padding: 32px 20px 64px;
    }}
    .hero {{
      background: linear-gradient(135deg, rgba(196,106,47,0.95), rgba(131,65,26,0.94));
      color: white;
      border-radius: 28px;
      padding: 28px;
      box-shadow: var(--shadow);
    }}
    .hero h1 {{
      margin: 0 0 10px;
      font-size: clamp(2rem, 5vw, 3.5rem);
      line-height: 1.05;
    }}
    .hero p {{
      margin: 6px 0;
      font-size: 1rem;
      color: rgba(255,255,255,0.88);
    }}
    .actions {{
      margin-top: 18px;
      display: flex;
      gap: 12px;
      flex-wrap: wrap;
    }}
    .button {{
      appearance: none;
      border: none;
      background: white;
      color: var(--accent-dark);
      padding: 12px 18px;
      border-radius: 999px;
      font-weight: 700;
      cursor: pointer;
    }}
    .button.secondary {{
      background: rgba(255,255,255,0.12);
      color: white;
      border: 1px solid rgba(255,255,255,0.25);
    }}
    .kpi-grid {{
      margin: 24px 0 12px;
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
      gap: 16px;
    }}
    .kpi-card, .section {{
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 24px;
      box-shadow: var(--shadow);
    }}
    .kpi-card {{
      padding: 20px;
    }}
    .kpi-label {{
      margin: 0 0 10px;
      color: var(--muted);
      text-transform: uppercase;
      letter-spacing: 0.08em;
      font-size: 0.78rem;
    }}
    .kpi-value {{
      margin: 0;
      font-size: 2.1rem;
      font-weight: 700;
    }}
    .kpi-note {{
      margin: 10px 0 0;
      color: var(--muted);
      font-size: 0.95rem;
    }}
    .section {{
      margin-top: 18px;
      padding: 22px;
    }}
    .section h2 {{
      margin: 0 0 8px;
      font-size: 1.45rem;
    }}
    .section p {{
      color: var(--muted);
      margin: 0 0 18px;
    }}
    .chart-grid {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(320px, 1fr));
      gap: 16px;
    }}
    .chart-card {{
      background: #fff;
      border: 1px solid var(--line);
      border-radius: 18px;
      padding: 16px;
    }}
    .chart-card h3 {{
      margin: 0 0 12px;
      font-size: 1rem;
    }}
    .chart-card img {{
      width: 100%;
      height: auto;
      display: block;
      border-radius: 14px;
      background: #fff;
    }}
    .data-table {{
      width: 100%;
      border-collapse: collapse;
      table-layout: fixed;
      font-size: 0.95rem;
    }}
    .data-table th,
    .data-table td {{
      padding: 12px 10px;
      text-align: left;
      border-bottom: 1px solid var(--line);
      vertical-align: top;
      white-space: normal;
      overflow-wrap: anywhere;
      word-break: break-word;
    }}
    .data-table th {{
      color: var(--muted);
      font-size: 0.82rem;
      text-transform: uppercase;
      letter-spacing: 0.06em;
    }}
    .table-wrap {{
      border: 1px solid var(--line);
      border-radius: 18px;
      background: white;
    }}
    .table-wrap .data-table th:first-child,
    .table-wrap .data-table td:first-child {{
      width: 28%;
    }}
    .table-wrap .data-table th:nth-child(2),
    .table-wrap .data-table td:nth-child(2) {{
      width: 14%;
    }}
    .table-wrap .data-table th:nth-child(3),
    .table-wrap .data-table td:nth-child(3),
    .table-wrap .data-table th:nth-child(4),
    .table-wrap .data-table td:nth-child(4),
    .table-wrap .data-table th:nth-child(5),
    .table-wrap .data-table td:nth-child(5) {{
      width: 14%;
    }}
    .table-wrap .data-table th:last-child,
    .table-wrap .data-table td:last-child {{
      width: 16%;
    }}
    .observation-grid {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
      gap: 16px;
    }}
    .observation-card {{
      border: 1px solid var(--line);
      border-radius: 20px;
      padding: 18px;
      background: white;
    }}
    .observation-card.attention {{
      background: linear-gradient(180deg, var(--soft-red), #fff 34%);
    }}
    .observation-card.positive {{
      background: linear-gradient(180deg, var(--soft-green), #fff 34%);
    }}
    .obs-header {{
      display: flex;
      justify-content: space-between;
      gap: 12px;
      align-items: flex-start;
      margin-bottom: 14px;
    }}
    .obs-header h3 {{
      margin: 0;
      font-size: 1.05rem;
    }}
    .obs-header p {{
      margin: 6px 0 0;
      font-size: 0.92rem;
    }}
    .obs-status {{
      white-space: nowrap;
      padding: 8px 12px;
      border-radius: 999px;
      background: rgba(32,32,32,0.06);
      font-size: 0.82rem;
      font-weight: 700;
    }}
    .observation-list {{
      margin: 0;
      padding-left: 0;
      list-style: none;
    }}
    .obs-item {{
      margin-bottom: 10px;
      line-height: 1.5;
      padding: 12px;
      border-radius: 14px;
      border: 1px solid var(--line);
      background: rgba(255,255,255,0.72);
    }}
    .obs-item p {{
      margin: 8px 0 0;
      color: var(--ink);
    }}
    .obs-meta {{
      display: flex;
      align-items: center;
      gap: 8px;
      flex-wrap: wrap;
    }}
    .obs-item.attention {{
      border-color: #e7c4b9;
    }}
    .obs-item.positive {{
      border-color: #c9ddcb;
    }}
    .obs-emoji {{
      font-size: 1rem;
      line-height: 1;
    }}
    .obs-pill {{
      display: inline-block;
      padding: 5px 10px;
      border-radius: 999px;
      font-size: 0.78rem;
      font-weight: 700;
      letter-spacing: 0.02em;
      background: var(--soft-gold);
      color: #724711;
    }}
    .obs-ticket {{
      display: inline-block;
      padding: 4px 8px;
      border-radius: 999px;
      background: rgba(196,106,47,0.12);
      color: var(--accent-dark);
      font-size: 0.76rem;
      font-weight: 700;
    }}
    .obs-channel {{
      display: inline-block;
      padding: 4px 8px;
      border-radius: 999px;
      background: rgba(32,32,32,0.06);
      color: var(--muted);
      font-size: 0.76rem;
      font-weight: 700;
    }}
    .footer {{
      color: var(--muted);
      margin-top: 22px;
      text-align: right;
      font-style: italic;
    }}
    .empty-state {{
      color: var(--muted);
      font-style: italic;
    }}
    @media print {{
      body {{
        background: white;
      }}
      .page {{
        max-width: none;
        padding: 0;
      }}
      .hero, .section, .kpi-card, .chart-card {{
        box-shadow: none;
      }}
      .actions {{
        display: none;
      }}
      .section, .kpi-card {{
        break-inside: avoid;
      }}
      .table-wrap {{
        overflow: visible;
        border-radius: 0;
      }}
      .data-table {{
        width: 100%;
        table-layout: fixed;
        font-size: 0.78rem;
      }}
      .data-table th,
      .data-table td {{
        padding: 7px 6px;
        white-space: normal;
        overflow-wrap: anywhere;
        word-break: break-word;
      }}
    }}
  </style>
</head>
<body>
  <main class="page">
    <section class="hero">
      <p>Relatório visual de desempenho</p>
      <h1>{html.escape(display_free_text(title))}</h1>
      <p><strong>Período:</strong> {html.escape(period)}</p>
      <p><strong>Data de emissão:</strong> {html.escape(emission_date)}</p>
      <div class="actions">
        <button class="button" onclick="window.print()">Salvar em PDF</button>
        <button class="button secondary" onclick="window.scrollTo({{top: document.body.scrollHeight, behavior: 'smooth'}})">Ver observações</button>
      </div>
    </section>

    <section class="kpi-grid">
      {kpi_cards_html}
    </section>

    <section class="section">
      <h2>Desempenho Consolidado</h2>
      <p>Visão geral dos critérios avaliados no período analisado.</p>
      <div class="chart-grid">
        <article class="chart-card">
          <h3>Análise Consolidada de Critérios</h3>
          <img src="{consolidated_chart_uri}" alt="Gráfico de análise consolidada de critérios">
        </article>
        <article class="chart-card">
          <h3>Desempenho Geral por Colaborador</h3>
          <img src="{collaborator_chart_uri}" alt="Gráfico de desempenho por colaborador">
        </article>
      </div>
    </section>

    <section class="section">
      <h2>Tabelas Analíticas</h2>
      <p>Detalhamento dos resultados consolidados e ranking individual, com foco nos indicadores principais.</p>
      <div class="table-wrap">
      {dataframe_to_html_table(consolidated_df)}
      </div>
      <div style="height: 18px;"></div>
      <div class="table-wrap">
      {dataframe_to_html_table(collaborator_table_df, float_cols={"Taxa de Cumprimento Geral"})}
      </div>
    </section>

    <section class="section">
      <h2>Observações Detalhadas</h2>
      <p>Todos os acolhedores em ordem alfabética, com todos os tickets atendidos e indicação visual de atendimento correto ou que exige atenção.</p>
      <div class="observation-grid">
        {observations_html}
      </div>
    </section>

    <section class="section">
      <p class="footer">Atenciosamente, Hemily Marques</p>
    </section>
  </main>
</body>
</html>
"""


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Gera relatorio de atendimentos a partir de CSV.")
    parser.add_argument("--input", required=True, help="Caminho do arquivo CSV de entrada.")
    parser.add_argument("--periodo", help="Periodo exibido no titulo do relatorio. Se omitido, sera inferido da coluna de data do CSV.")
    parser.add_argument("--titulo", default="Relatorio Semanal de Atendimentos Online", help="Titulo principal.")
    parser.add_argument("--encoding", default="utf-8", help="Codificacao do CSV.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    base_path = Path(__file__).resolve().parent
    _, output_dir = ensure_dirs(base_path)
    os.environ.setdefault("MPLCONFIGDIR", str(base_path / ".mplconfig"))

    import matplotlib

    matplotlib.use("Agg")

    global plt
    import matplotlib.pyplot as plt

    csv_path = Path(args.input).resolve()
    df = load_dataframe(csv_path, args.encoding)
    prepared_df, _ = prepare_dataframe(df)
    inferred_period = infer_period_from_dataframe(df)
    report_period = inferred_period or args.periodo or "Periodo nao identificado"

    consolidated = [
        summarize_binary_criterion(prepared_df, "__obs_interna__", "Observacao Interna"),
        summarize_binary_criterion(prepared_df, "__resolucao__", "Resolucao"),
        summarize_binary_criterion(prepared_df, "__zendesk__", "Cadastro Zendesk"),
        summarize_general_evaluation(prepared_df),
    ]

    collaborator_df = build_collaborator_table(prepared_df)
    channel_df = build_channel_summary(prepared_df)
    observation_groups = build_observation_groups(prepared_df)

    consolidated_chart = output_dir / "analise_criterios.png"
    collaborator_chart = output_dir / "desempenho_colaboradores.png"
    plot_consolidated_chart(consolidated, consolidated_chart)
    plot_collaborator_chart(collaborator_df, collaborator_chart)

    report_markdown = generate_markdown_report(
        period=report_period,
        title=args.titulo,
        consolidated=consolidated,
        collaborator_df=collaborator_df,
        observation_groups=observation_groups,
        output_dir=output_dir,
    )

    report_path = output_dir / "relatorio_atendimentos.md"
    report_path.write_text(report_markdown, encoding="utf-8")
    report_html = generate_html_report(
        period=report_period,
        title=args.titulo,
        consolidated=consolidated,
        collaborator_df=collaborator_df,
        channel_df=channel_df,
        observation_groups=observation_groups,
        output_dir=output_dir,
    )
    html_path = output_dir / "index.html"
    html_path.write_text(report_html, encoding="utf-8")
    print(f"Relatorio gerado em: {report_path}")
    print(f"Relatorio visual gerado em: {html_path}")


if __name__ == "__main__":
    main()
