# FolderSuggest

An intelligent Outlook add-in that automatically suggests folder classifications for emails using machine learning. Uses ML.NET to train a multiclass classifier on your email folder structure and predicts the best folder for each email.

## Features

- **Automatic Folder Prediction**: When you select an email in Outlook, the add-in predicts the top 3 most likely folders
- **Confidence Scores**: Each prediction includes a confidence percentage
- **One-Click Filing**: Click a predicted folder button to instantly move the email
- **Model Training**: Built-in training UI to create a custom model from your existing emails
- **Multi-Explorer Support**: Works seamlessly with multiple Outlook windows

## Architecture

### Components

1. **FolderPredictionTrainer** - Trains an ML.NET multiclass classification model
   - Collects emails from Outlook folders
   - Extracts features from Subject and SenderEmail
   - Saves model as `EmailFolderModel.zip` and labels as `FolderLabels.json`
   - Uses SdcaMaximumEntropy trainer with text featurization

2. **FolderPredictionService** - Loads and uses the trained model for inference
   - Loads model and label mapping at startup
   - Predicts folder for given email subject and sender
   - Returns top N predictions with confidence scores

3. **ThisAddIn** - Main VSTO add-in entry point
   - Manages Outlook event subscriptions
   - Hooks SelectionChange events on all explorer windows
   - Triggers predictions when emails are selected
   - Updates ribbon buttons with predictions

4. **Ribbon1** - Custom ribbon UI
   - Displays top 3 predicted folders as buttons (Folder1-3)
   - Shows confidence scores (e.g., "Legal - Misc (87%)")
   - Handles folder button clicks to move emails
   - Training button (Folder5) opens training dialog

## Getting Started

### Requirements

- Visual Studio 2022
- Office 2016 or later (with Outlook)
- .NET Framework 4.8
- NuGet packages: Microsoft.ML 3.0.0, Microsoft.ML.FastTree 3.0.0

### Installation

1. Clone the repository
2. Open `FolderSuggest.sln` in Visual Studio
3. Build the solution (Debug or Release)
4. Press F5 to deploy the add-in

### Usage

#### Training a Model

1. Organize your emails into folders (the model learns from folder structure)
2. In Outlook, go to the **FolderSuggest** tab
3. Click **Training** button
4. Click **Start Training**
5. Wait for the process to complete (shows progress percentage)

The model will be saved to `%APPDATA%\FolderSuggest\`:
- `EmailFolderModel.zip` - The trained ML.NET model
- `FolderLabels.json` - List of folder names in order

#### Using Predictions

1. Click on any email in Outlook
2. Check the **FolderSuggest** tab - Folder1-3 buttons update with predicted folders and confidence scores
3. Click a button to move the email to that folder
4. If no model exists, buttons appear disabled

#### Training Data Collection

The trainer automatically:
- Scans all folders under Inbox (except root Inbox)
- Extracts email Subject and SenderEmail fields
- Creates a multiclass dataset labeled by folder name

## File Structure

```
FolderSuggest/
├── Models/
│   ├── EmailItem.cs           # Display model for UI
│   └── EmailData.cs           # ML training model + prediction output
├── Services/
│   ├── FolderPredictionTrainer.cs      # Training pipeline
│   ├── FolderPredictionService.cs      # Inference engine
│   └── OutlookService.cs      # Outlook data collection
├── ViewModels/
│   ├── EmailListViewModel.cs  # Email list UI logic
│   └── TrainingViewModel.cs   # Training progress UI logic
├── Views/
│   ├── EmailListWindow.xaml   # Email list view
│   └── TrainingWindow.xaml    # Training progress dialog
├── ThisAddIn.cs               # Add-in entry point & event handling
├── Ribbon1.cs                 # Ribbon UI logic & predictions
└── Ribbon1.xml                # Ribbon UI layout
```

## How It Works

### Training Pipeline

1. **Data Collection**: Recursively collects emails from folders, extracting Subject and SenderEmail
2. **Feature Extraction**: 
   - Text featurization on Subject → SubjectFeatures
   - Text featurization on SenderEmail → SenderFeatures
   - Concatenates both into Features vector
3. **Label Encoding**: Maps folder names to numeric labels via MapValueToKey
4. **Model Training**: Uses SdcaMaximumEntropy (Stochastic Dual Coordinate Ascent) for multiclass classification
5. **Label Mapping**: Saves folder names to JSON file to reverse the label encoding
6. **Model Serialization**: Saves trained model to zip file

### Prediction Pipeline

1. **Event Hook**: SelectionChange event fires when user clicks an email
2. **Feature Extraction**: Extracts Subject and SenderEmail from selected MailItem
3. **Prediction**: Runs through ML model to get probability scores for each folder
4. **Mapping**: Maps score array indices to folder names using saved labels
5. **UI Update**: Updates ribbon buttons with top 3 predictions and scores
6. **User Action**: User clicks a button to move email via MailItem.Move()

## Debug Output

Extensive debug logging is available in Visual Studio's Output window when running in debug mode. Look for:

```
=== FolderSuggest AddIn Starting ===
Prediction service loaded: True
Successfully hooked explorer
OnSelectionChange called
Selection count: 1
Processing email: [Subject]
Predictions received: 3
```

## Known Limitations

- Model training only uses emails already filed in subfolders (root Inbox items are skipped)
- Predictions are based only on Subject and SenderEmail (future: could add Body, Date, etc.)
- Folder structure must be consistent between training and prediction

## Future Enhancements

- [ ] Add Body text as training feature
- [ ] Support for different Outlook folder types (Tasks, Contacts, etc.)
- [ ] Retraining UI to improve model
- [ ] Performance metrics on trained model
- [ ] Confidence threshold setting
- [ ] Undo functionality for moved emails

## License

MIT

## Contributing

Contributions welcome! Please feel free to submit pull requests.
