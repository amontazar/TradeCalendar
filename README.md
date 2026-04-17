# Trade Calendar AddOn (NinjaTrader 8)

This repository contains a NinjaTrader 8 AddOn that provides a calendar-style realized PnL view with filtering and CSV import/export.

## Prerequisites

- NinjaTrader 8 installed
- Access to your NinjaTrader user folder:
  - `Documents\NinjaTrader 8\bin\Custom\AddOns`

## Install

1. Copy `TradeCalendarAddOn.cs` into:
   - `Documents\NinjaTrader 8\bin\Custom\AddOns\TradeCalendar\`
2. Open NinjaTrader 8.
3. Go to **New > NinjaScript Editor**.
4. In the editor, click **Compile** (F5).
5. Wait for compile to complete with no errors.

## Open the AddOn

1. In NinjaTrader Control Center, go to **New**.
2. Click **Trade Calendar**.

## Update to a New Version

1. Close the Trade Calendar window (and preferably close NinjaTrader).
2. Replace `TradeCalendarAddOn.cs` with the updated file.
3. Reopen NinjaTrader.
4. Compile again from **NinjaScript Editor** (F5).

## Data and Settings Files

The AddOn stores ledgers/UI state in NinjaTrader user data files:

- `TradeCalendarLedger.xml`
- `TradeCalendarImportedTrades.xml`
- `TradeCalendarUiState.xml`

These are created under your NinjaTrader user data directory.

## Troubleshooting

- AddOn not visible under **New**:
  - Recompile in NinjaScript Editor.
  - Check compile errors in the editor output.
- Old behavior still showing:
  - Restart NinjaTrader after compiling.
- Account/default selection not as expected:
  - Close and reopen the AddOn after changing account.
  - If needed, restart NinjaTrader to refresh selector state.
