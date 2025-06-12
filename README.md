# MCExcel - Monte Carlo Reading Simulation

This Excel workbook is a hands-on demo of how Monte Carlo simulation can forecast book completion times based on how much you’ve actually been reading day to day: [backstory of this project](https://medium.com/thrivve-partners/i-asked-a-room-full-of-leaders-to-predict-the-future-with-dice-076618d028ab)

## Overview

This workbook models the uncertainty in daily reading productivity by using historical throughput data to generate probabilistic forecasts for completing a book. It runs multiple simulations to provide statistical insights into likely completion timeframes.

## What's Inside

### 📊 Sheets Overview

1. **Your Results** - Summary of simulation outcomes with frequency analysis
2. **Graph** - Visualization data and charts for result distribution
3. **Simulation** - Core Monte Carlo engine with 1,000 parallel simulations
4. **Throughput** - Historical daily reading data (pages per day)

### 🎯 Key Features

- **1,000 parallel simulations** running simultaneously
- **Dynamic lookup table** mapping random values to actual reading throughput
- **Statistical analysis** including percentiles, min/max values
- **Visual distribution** of probable outcomes
- **Configurable parameters** for different book lengths
- **Variable start date** to model different outcomes based on date

## How It Works

### 1. Data Collection
The **Throughput** sheet contains historical reading data:
- Date-stamped daily reading sessions
- Pages read per day (ranging from 0-22 pages)
- Represents real reading variability, including "off days"

### 2. Probability Distribution
The simulation uses a lookup table that maps random dice rolls (1-20) to actual reading throughput values observed in your historical data:

| Roll | Pages | Roll | Pages | Roll | Pages |
|------|-------|------|-------|------|-------|
| 1    | 4     | 8    | 14    | 15   | 8     |
| 2    | 6     | 9    | 8     | 16   | 0     |
| 3    | 7     | 10   | 3     | 17   | 12    |
| 4    | 0     | 11   | 4     | 18   | 9     |
| 5    | 8     | 12   | 9     | 19   | 4     |
| 6    | 22    | 13   | 3     | 20   | 11    |
| 7    | 0     | 14   | 7     | —    | —     |

### 3. Monte Carlo Simulation
- **Starting point**: 164 pages remaining
- **Method**: Each simulation runs independently, rolling random numbers daily
- **Process**: Subtract daily pages read until book completion
- **Output**: Days required to finish for each simulation

### 4. Statistical Analysis
The workbook provides comprehensive statistics:
- **50th percentile (Median)**: 24 days
- **70th percentile**: 26 days  
- **85th percentile**: 28 days
- **95th percentile**: 30 days
- **Fastest completion**: 14 days
- **Slowest completion**: 37 days

## Using the Workbook

### Quick Start
1. Open the workbook in Excel
2. View **Your Results** sheet for summary statistics
3. Check **Graph** sheet for visual distribution
4. Press F9 to recalculate and run new simulations

### Customization
To adapt for your own reading project:

1. **Update Throughput Data**:
   - Replace data in the **Throughput** sheet with your actual reading history
   - Include both productive and unproductive days for accuracy

2. **Modify Book Length**:
   - Change the "Pages Left" value in cell D2 of the **Simulation** sheet
   - The simulation will automatically recalculate

3. **Adjust Lookup Table**:
   - Update the Roll Value → Pages Read mapping in columns A-B of **Simulation** sheet
   - Ensure it reflects your actual reading patterns

### Understanding the Results

**Percentile Interpretation**:
- 50th percentile: You have a 50% chance of finishing by this day
- 85th percentile: You have an 85% chance of finishing by this day
- 95th percentile: Conservative estimate with high confidence

## Distribution Insights:
- Key observation: Even the most common result occurs in only a small fraction of simulations
- This demonstrates the inherent unpredictability in day-to-day reading patterns (and any throughput systems, e.g. software delivery)
- No single outcome dominates - results are spread across a wide range of possibilities
- The "typical" completion time is actually quite rare when viewed individually

**Range Analysis**:
- The difference between fastest (14 days) and slowest (37 days) shows the impact of reading consistency
- Zero-page days significantly affect completion time

## Technical Details

### Simulation Engine
- **Random Number Generation**: Uses Excel's RAND() function
- **Lookup Method**: INDEX/MATCH functions for throughput mapping  
- **Iteration**: Daily simulation until pages remaining ≤ 0
- **Scale**: 1,000 independent simulation runs

### Data Structure
- **Simulation Matrix**: 1,000 columns × 50 rows (max days)
- **Results Row**: Row 53 contains completion days for each simulation
- **Statistics**: Calculated using PERCENTILE, MIN, MAX functions

## Applications

This example model is useful for:
- **Project planning**: Realistic timeline estimation
- **Goal setting**: Understanding probability of meeting deadlines
- **Risk assessment**: Identifying range of possible outcomes
- **Performance analysis**: Impact of consistency on results

## Prerequisites

- Microsoft Excel 2016 or later
- Basic understanding of probability concepts
- Familiarity with Excel functions (optional for customization)

## Tips for Best Results

1. **Include bad days**: Don't sanitize your throughput data - include zero and low-productivity days
2. **Sufficient history**: Use at least 2-3 weeks of reading data for stable patterns
3. **Regular updates**: Refresh simulations periodically as you gather more data
4. **Context matters**: Consider external factors that might affect future reading patterns

## File Information

- **Version**: 1.0 Public
- **Format**: Excel (.xlsx)
- **Compatibility**: Excel 2016+, Excel Online, Google Sheets (with limitations)
- **Size**: Contains extensive simulation data and formulas

---

*This workbook demonstrates practical applications of Monte Carlo simulation for personal productivity forecasting. The techniques can be adapted for various project estimation scenarios beyond reading.*
