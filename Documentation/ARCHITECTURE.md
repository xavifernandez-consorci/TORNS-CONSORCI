# TORNS CONSORCI Architecture

## Overview
This project is structured as a modular Excel VBA scheduling application.
The architecture is intentionally separated into layers so configuration, domain rules, persistence, and UI can evolve independently.

## Layers
- Workbook layer: workbook lifecycle and startup coordination.
- Application layer: bootstrap and high-level orchestration.
- Domain layer: configuration, employees, schedule context, and assignments.
- Service layer: rotation motor and duty services.
- Data layer: loading and saving task data from workbook storage.
- Presentation layer: forms for configuration, main navigation, and schedule review.

## Responsibilities
- Configuration: hold all configurable settings and their persistence.
- Scheduling: generate and regenerate schedule states according to rules.
- Intensive rotation: manage the intensive cycle and related duty assignments.
- Duty management: manage primary and backup duty operators.
- Data access: isolate worksheet interactions from the remainder of the application.

## Design Principles
- SOLID principles applied at module and class boundaries.
- No hardcoded business values in core modules.
- No business logic implemented in this architecture scaffold.
- VBA modules are documented and intended to remain modular.
