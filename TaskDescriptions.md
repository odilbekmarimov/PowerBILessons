# Factory Dashboard Projects

## Project 1: Basic Factory Performance Dashboard

### Technical Description
In this initial project, students will create a single-page dashboard focusing on core factory performance metrics. Using a manufacturing dataset containing production volumes, quality metrics, and basic operational data, students will:

### Data Import & Preparation
- Import factory production data from Excel/CSV files containing daily production logs
- Clean inconsistent date formats and standardize product names
- Create a simple data model connecting production and defect tables

### Key Visualizations
- Production volume by product line (bar chart)
- Daily production trend (line chart)
- Quality metrics - % defective by product (column chart)
- Production KPI cards showing total units, defect rate, and on-time delivery %
- Simple slicer panel for filtering by date range and product line

### Technical Requirements
- Implement basic data transformations in Power Query
- Create relationships between production and product tables
- Use built-in aggregations (SUM, AVERAGE, COUNT)
- Apply basic formatting and consistent color schemes
- Implement page-level filters

### Deliverable
A single-page operational dashboard that factory floor managers could use to monitor daily production performance and quality metrics.

---

## Project 2: Sales Performance Dashboard with DAX

### Technical Description
Building on Project 1, students will now create a multi-page dashboard that expands the analysis to include sales data and financial performance using DAX measures. The project will connect production data with sales outcomes.

### Data Model Enhancement
- Import additional tables: sales transactions, customer data, product costs
- Implement proper star schema modeling with dimension and fact tables
- Create a date dimension table using DAX
- Establish appropriate relationships between production, inventory, and sales tables

### DAX Implementations
- YTD Sales = TOTALYTD(SUM(Sales[Amount]), Dates[Date])
- MoM Growth % = DIVIDE(SUM(Sales[Amount]) - [Previous Month Sales], [Previous Month Sales])
- Production Efficiency = DIVIDE(SUM(Production[Units]), SUM(Production[Labor_Hours]))
- Profit Margin = DIVIDE(SUM(Sales[Amount]) - SUM(Production[Cost]), SUM(Sales[Amount]))
- CALCULATE-based measures for comparing current performance vs targets

### Advanced Visualizations
- Sales vs. Production chart showing alignment between manufacturing and market demand
- YTD sales performance with comparison to previous year
- Production efficiency by manufacturing line with trend analysis
- Profitability analysis by product category and customer segment
- Time-based charts with drill-down capabilities from year to quarter to month

### Technical Requirements
- Proper implementation of filter context in DAX measures
- Creation of calculated columns for categorization
- Implementation of time intelligence functions
- Use of variables in complex DAX formulas
- Drill-through functionality from summary to detail

### Deliverable
A 3-page interactive dashboard connecting production metrics to sales performance and financial outcomes, suitable for middle management reporting.

---

## Project 3: Advanced Factory Analytics Dashboard

### Technical Description
In this comprehensive final project, students will develop a complete business intelligence solution for factory management, incorporating advanced DAX, complex visuals, and predictive elements. The project simulates a real-world business intelligence deployment for manufacturing optimization.

### Complex Data Integration
- Connect to multiple data sources including SQL database for production data, Excel files for financial data, and CSV exports from machinery sensors
- Implement incremental refresh patterns for large historical datasets
- Create composite models connecting direct query and import mode tables
- Implement row-level security based on department roles (production, sales, executive)

### Advanced DAX Modeling
- Create dynamic rolling calculations for variable time periods
- Implement production forecasting measures using time series functions
- Build complex KPIs combining multiple business areas:
  - Overall Equipment Effectiveness (OEE) = CALCULATE(PRODUCT(DIVIDE([Availability], 100), DIVIDE([Performance], 100), DIVIDE([Quality], 100)))
  - Cost Per Unit = DIVIDE(SUM(Costs[Material]) + SUM(Costs[Labor]) + SUM(Costs[Overhead]), SUM(Production[Units]))
  - Cash Conversion Cycle measures connecting production lead time to customer payment timelines
- Implement DAX patterns for anomaly detection in production metrics

### Advanced Visualization Techniques
- Custom visual layout with navigation buttons and bookmarks for different analysis paths
- Manufacturing process flow diagram with performance indicators at each stage
- Small multiples visualization for comparing performance across production lines
- Waterfall charts showing cost breakdown and profitability drivers
- Map visualization showing factory locations with supply chain metrics
- Decomposition tree for root cause analysis of production issues
- What-if parameters for production capacity planning

### Technical Requirements
- Advanced DAX using iteration functions (SUMX, FILTER)
- Implementation of variables for optimal performance
- Proper use of context transition
- Creation of parameter tables for scenario analysis
- Implementation of calculation groups for switching metrics
- Optimization for performance using appropriate DAX patterns
- Report-level tooltips with detailed analysis
- Mobile-optimized view for executive dashboards

### Deliverable
A comprehensive multi-page business intelligence solution that connects production, quality, financial, and sales data into an integrated view of factory performance, with advanced analytics capabilities suitable for executive decision-making.
