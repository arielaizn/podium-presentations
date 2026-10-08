# Podium Presentations

A presentation-creation skill adapted from [Guy Aga’s Podium](https://github.com/guyaga/10d10s-day09-presentations), with Hebrew/English workflows, RTL support, and a coordinated presentation package.

## Default deliverables

- Editable PowerPoint (`.pptx`)
- Matching static presentation PDF
- Animated HTML presentation, with a verified Vercel or Here.now link when authorized publishing is available
- At least six distinct, topic-specific companion PDF documents

A later explicit request can narrow this scope. Publishing and external services remain subject to the host assistant’s permissions and available tools; this repository does not provide hosting credentials.

## Install and use

Copy the repository’s contents into a `podium` directory in your assistant’s supported skills location, preserving its subdirectories. Use that platform’s skill installation instructions. The entry point is [SKILL.md](SKILL.md); do not copy private credentials into the skill.

Ask the assistant to create a presentation and supply the topic, audience, goal, language, and any brand/source material. The skill defines the output contract, research and design workflow, companion documents, and verification checks.

## Runtime requirements

The bundled Python engine requires Python 3.10+, `python-pptx`, Pillow, lxml, and fonttools. PyMuPDF is required by the package checker. Linux PDF rendering uses LibreOffice and PyMuPDF or pdftoppm. ffmpeg/ffprobe are optional for video. Hebrew-capable fonts are required for Hebrew output.

Run `python scripts/engine/doctor.py` to inspect dependencies. Platform-required presentation authoring tools take precedence over the bundled engine.

## Verification scope

The adapted package was checked with Hebrew PPTX/PDF rendering and visual review, six companion-PDF structural checks, Hebrew companion-PDF rendering and visual review, and JavaScript navigation behavior tests. Live-browser visual QA of the animated HTML presentation was unavailable in the test environment. Publication of an individual presentation must be verified when that presentation is produced; no presentation-hosting deployment is included here.

## Attribution and license status

See [provenance](references/provenance.md) for retained upstream files, adaptations, and the pinned source commit. No license file was found in the inspected upstream tree. This repository does not assert an open-source license or grant rights to upstream material. Preserve attribution and obtain any necessary rights-holder permission before redistribution or reuse that requires it.
