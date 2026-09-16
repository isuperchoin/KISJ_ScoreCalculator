# KISJ Score Calculator

> An automated GPA and grade calculation tool designed for students at Korea International School Jeju (KISJ).

![Python](https://img.shields.io/badge/Python-3.8+-3776AB?style=flat-square&logo=python&logoColor=white)
![BuiltWith](https://img.shields.io/badge/Built%20With-Claude%20Code-7C3AED?style=flat-square)
![License](https://img.shields.io/badge/License-MIT-green.svg?style=flat-square)

---

# Overview

This school utility automates the process of calculating course grades and cumulative GPAs in accordance with school grading scales.

The underlying calculation logic was structured based on official school syllabus guidelines, leveraging Prompt Engineering and AI-Assisted Development (Claude Code) to prototype, refine, and test the implementation.

---

# Key Features

- **GPA/Letter Grade Calculation**: Calculates term and cumulative GPAs based on the official KISJ letter grade policy (A+ 98-100 ... D- 60-62, F 50-59, NG below 50).
- **Overwrite Policy**: Tick "Apply Overwrite Policy" on any formative score and pair it with a summative; when the summative is higher, the formative counts as that score. Pairings are kept in exported reports and restored on import.
- **Result Exporting System**: Allows users to export the result as PDF, JPG, Excel, and CSV.
- **Customizable User Input**: Allows users to customize Quarters, Subjects, The number of assesment, etc to best fit their current situation.
- **Importing Previous Result**: Allows users to import their previous result so they don't have to enter every scores every time.
- **Community Deployment**: Designed for lightweight distribution and usage within the school student community.

---

# Tech Stack & Development Workflow

- **Language**: Python 3.x
- **Development Process**: 
  - **System Design & Requirements**: Choin (Logic Definition, Edge Case Mapping, Testing)
  - **Code Generation & Refinement**: Claude Code (AI CLI Tool)
- **Version Control**: Git / GitHub

---

# How to Run

### Installation

1. Go to the [Releases](https://github.com/isuperchoin/KISJ_ScoreCalculator/releases) page on this repository.
2. Download the latest `KISJ_ScoreCalculator.app.zip` file.
3. Unzip the file and double-click `KISJ_ScoreCalculator.app` to launch.

> **Note on macOS Security Warning (Gatekeeper)**:
> Since this app is not signed with an Apple Developer Certificate, macOS may display a security prompt: *"App cannot be opened because it is from an unidentified developer."*
> - **Solution**: Right-click (or `Control` + click) the `.app` file ---> select **Open** ---> click **Open** in the confirmation dialog.
(This is only required on the first launch.)
