# PowerPoint Automate

A tool to automate the generation of the Prizes PowerPoint.

## Usage

In order to use this tool, you need to do the following:

### Folder Structure

Before running this tool, it requires you having a directory with the following folder structure:

```
Prizes Generator/
├─ Logos/
│  ├─ Logo1.png
│  ├─ ...
├─ Participants/
│  ├─ StandPhoto1.png
│  ├─ ...
├─ Prizes.csv
├─ TemplatePresentation.pptx
```

> [!WARNING]  
> Be aware that the `Logos` and `Participants` need to have that exact name or else the program won't work.

### CSV File

| PrizeName             | PrizeLogo | Title         | ParticipantImage | Author1     | Author2 | Author3 | Author4 |
|-----------------------|-----------|---------------|------------------|-------------|---------|---------|---------|
| The name of the prize | Logo1.png | Project title | StandPhoto1.png  | Author Name |         |         |         |

_NOTE: If a prize doesn't have any winner, leave the title, participant and authors empty and the generator will mark
them as so._

> [!WARNING]  
> Pay attention to the CSV header names, as they must be the exact same.
