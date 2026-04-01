# Proposta de Dashboard Complementar

## Objetivo

Complementar o relatorio estatico com uma visualizacao interativa para acompanhamento de desempenho da equipe, filtros por periodo e consulta rapida por colaborador.

## KPIs principais

- Taxa de cumprimento geral da equipe
- Cumprimento por criterio
- Ranking de colaboradores
- Total de atendimentos por canal
- Distribuicao da avaliacao geral
- Tendencia historica por semana ou mes

## Layout sugerido

### 1. Filtros globais

- Periodo
- Colaborador
- Tipo de atendimento

### 2. Cartoes de KPI

- Taxa de cumprimento geral
- Total de atendimentos
- Percentual de avaliacao correta
- Percentual de registros com atencao

### 3. Desempenho consolidado

- Grafico de barras para cumprimento por criterio
- Grafico donut para avaliacao geral

### 4. Desempenho individual

- Ranking horizontal por colaborador
- Tabela detalhada com filtros e busca

### 5. Observacoes

- Lista filtravel de observacoes com destaque visual para atencao

## Stack recomendada

### Backend

- Python
- pandas
- FastAPI

### Frontend

- React com TypeScript
- Plotly.js ou Chart.js
- Tailwind CSS
- TanStack Table

### Persistencia opcional

- PostgreSQL para historico e tendencias

## Endpoints sugeridos

- `GET /api/kpis`
- `GET /api/criterios`
- `GET /api/colaboradores`
- `GET /api/observacoes`
- `GET /api/tendencias`

## Modelo de entrega

### Fase 1

- Script Python gerando JSON consolidado a partir do CSV
- API simples servindo os indicadores

### Fase 2

- Dashboard web com filtros e rankings

### Fase 3

- Persistencia historica e acompanhamento de tendencias
