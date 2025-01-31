function generateCombinations() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var promptsSheet = ss.getSheetByName("Prompts");
  var valuesSheet = ss.getSheetByName("Values");
  var modelsSheet = ss.getSheetByName("Models");
  var parametersSheet = ss.getSheetByName("Parameters");
  var outputSheet = ss.getSheetByName("Output") || ss.insertSheet("Output");

  // Run validation before proceeding
  var validationMessage = validateData(promptsSheet, valuesSheet, modelsSheet, parametersSheet);
  if (validationMessage) {
    SpreadsheetApp.getUi().alert("Validation Error:\n\n" + validationMessage);
    return; // Stop execution if there's an error
  }

  // Clear previous output
  outputSheet.clear();

  // Read prompts
  var promptsData = promptsSheet.getDataRange().getValues();
  var promptHeaders = promptsData.shift(); // Extract headers
  var promptIdIndex = promptHeaders.indexOf("prompt_id");

  // Read values
  var valuesData = valuesSheet.getDataRange().getValues();
  var valueHeaders = valuesData.shift(); // Extract headers
  var valueIdIndex = valueHeaders.indexOf("prompt_id");

  // Remove the "Name" column if it exists
  var excludeColumns = ["Name"];
  var filteredValueHeaders = valueHeaders.filter(header => !excludeColumns.includes(header));
  var filteredValueIndices = filteredValueHeaders.map(header => valueHeaders.indexOf(header));

  // Read models
  var modelsData = modelsSheet.getDataRange().getValues();
  var modelHeaders = modelsData.shift();
  var modelNameIndex = modelHeaders.indexOf("Name");
  var modelVersionIndex = modelHeaders.indexOf("Version");

  // Read parameters and expand combinations
  var parametersData = parametersSheet.getDataRange().getValues();
  var parameterNames = parametersData.map(row => row[0]); // Extract parameter names
  var parameterValues = parametersData.map(row => {
    let value = row[1];
    if (typeof value === "number") {
      return [String(value)];
    }
    return value.split(",").map(val => val.trim());
  });

  var parameterCombinations = cartesianProduct(parameterValues);

  // Group values by prompt_id
  var valuesMap = {};
  valuesData.forEach(row => {
    var promptId = row[valueIdIndex];
    if (!valuesMap[promptId]) {
      valuesMap[promptId] = [];
    }
    valuesMap[promptId].push(row);
  });

  // Prepare output headers
  var headers = ["index", "prompt_id", "model_name", "model_version", ...parameterNames, "filled_prompt"];
  outputSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Process prompts, models, parameters, and generate output
  var outputData = [];
  var index = 1;

  promptsData.forEach(promptRow => {
    var promptId = promptRow[promptIdIndex];
    var promptTemplate = promptRow.slice(1).join(" "); // Merge all non-ID fields into a single string

    if (valuesMap[promptId]) {
      valuesMap[promptId].forEach(valueRow => {
        modelsData.forEach(modelRow => {
          parameterCombinations.forEach(paramSet => {
            let filledPrompt = promptTemplate;

            filteredValueIndices.forEach(idx => {
              let header = valueHeaders[idx];
              filledPrompt = filledPrompt.replace(new RegExp(`{{${header}}}`, 'g'), valueRow[idx]);
            });

            outputData.push([
              index,
              promptId,
              modelRow[modelNameIndex],   // Model Name
              modelRow[modelVersionIndex], // Model Version
              ...paramSet,                 // Expanded Parameter Values
              filledPrompt
            ]);
            index++;
          });
        });
      });
    }
  });

  // Write output to "Output" sheet
  if (outputData.length > 0) {
    outputSheet.getRange(2, 1, outputData.length, headers.length).setValues(outputData);
  }
}

// 🚀 New Validation Function
function validateData(promptsSheet, valuesSheet, modelsSheet, parametersSheet) {
  var errors = [];

  // Check if sheets exist
  if (!promptsSheet || !valuesSheet || !modelsSheet || !parametersSheet) {
    errors.push("One or more required sheets (Prompts, Values, Models, Parameters) are missing.");
  }

  // Check if required columns exist in Prompts sheet
  var promptsHeaders = promptsSheet.getDataRange().getValues()[0];
  if (!promptsHeaders.includes("prompt_id")) {
    errors.push("The Prompts sheet must have a 'prompt_id' column.");
  }

  // Check if required columns exist in Values sheet
  var valuesHeaders = valuesSheet.getDataRange().getValues()[0];
  if (!valuesHeaders.includes("prompt_id")) {
    errors.push("The Values sheet must have a 'prompt_id' column.");
  }

  // Check if required columns exist in Models sheet
  var modelsHeaders = modelsSheet.getDataRange().getValues()[0];
  if (!modelsHeaders.includes("Name") || !modelsHeaders.includes("Version")) {
    errors.push("The Models sheet must have 'Name' and 'Version' columns.");
  }

  // Check if Parameters sheet has valid data
  var parametersData = parametersSheet.getDataRange().getValues();
  if (parametersData.length === 0) {
    errors.push("The Parameters sheet is empty.");
  } else {
    parametersData.forEach(row => {
      if (!row[0] || !row[1]) {
        errors.push("Each parameter in the Parameters sheet must have a name and values.");
      }
    });
  }

  // Return all errors as a message
  return errors.length > 0 ? errors.join("\n") : null;
}

// Helper function to compute Cartesian product of multiple arrays
function cartesianProduct(arr) {
  return arr.reduce((a, b) => a.flatMap(d => b.map(e => d.concat(e))), [[]]);
}
