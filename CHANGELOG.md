# Changelog

All notable changes will be documented in this file.  
Versioning follows simple tags: `v0.10`, `v0.11`, …

## v0.10
- Balance reconciliation working with corrected sign handling.
- Stable ingest for Chase statements; dashboard populates.

## Unreleased
- Split Deposits into subcategories: Payroll, Transfer In, Check Deposit, Return.
- Remove post-ingest dedupe in favor of load-once guard.
- Add unit tests for parsers/reconciliation.
- Expand parsers to additional institutions.
