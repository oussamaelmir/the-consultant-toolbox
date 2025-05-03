Office.onReady(() => {
  console.log("Office JS ready!");

  Office.actions.associate("emptyTextBoxes", emptyTextBoxes);
  Office.actions.associate("emptyEntireSlide", emptyEntireSlide);
  Office.actions.associate("equalizeSize", equalizeSize);
  Office.actions.associate("equalizeHeight", equalizeHeight);
  Office.actions.associate("equalizeWidth", equalizeWidth);
  Office.actions.associate("stackHorizontally", stackHorizontally);
  Office.actions.associate("stackVertically", stackVertically);
  Office.actions.associate("invertPositions", invertPositions);
  Office.actions.associate("removeMargins", removeMargins);
  Office.actions.associate("toggleTextWrap", toggleTextWrap);
  Office.actions.associate("copyPosition", copyPosition);
  Office.actions.associate("pastePosition", pastePosition);
  Office.actions.associate("copyDimensions", copyDimensions);
  Office.actions.associate("pasteDimensions", pasteDimensions);
  Office.actions.associate("autoFitShapeToText", autoFitShapeToText);
  Office.actions.associate("autoFitTextToShape", autoFitTextToShape);
  Office.actions.associate("insertHarvey0", insertHarvey0);
  Office.actions.associate("insertHarvey25", insertHarvey25);
  Office.actions.associate("insertHarvey50r", insertHarvey50r);
  Office.actions.associate("insertHarvey50l", insertHarvey50l);
  Office.actions.associate("insertHarvey75", insertHarvey75);
  Office.actions.associate("insertHarvey100", insertHarvey100);
  Office.actions.associate("insertGreenLight", insertGreenLight);
  Office.actions.associate("insertOrangeLight", insertOrangeLight);
  Office.actions.associate("insertRedLight", insertRedLight);
  Office.actions.associate("insertUpArrow", insertUpArrow);
  Office.actions.associate("insertDownArrow", insertDownArrow);
  Office.actions.associate("insertEqualsSign", insertEqualsSign);
  Office.actions.associate("createPostItNote", createPostItNote);
  Office.actions.associate("insertFootnote", insertFootnote);
  Office.actions.associate("createPreliminaryStamp", createPreliminaryStamp);
  Office.actions.associate("createExampleStamp", createExampleStamp);
  Office.actions.associate("createExecSumStamp", createExecSumStamp);
  Office.actions.associate("createDraftStamp", createDraftStamp);
});

export async function emptyTextBoxes(event: Office.AddinCommands.Event) {
  console.log("emptyTextBoxes button clicked!");

  try {
    await PowerPoint.run(async (context) => {
      const selectedShapes = context.presentation.getSelectedShapes();
      selectedShapes.load("items/textFrame/hasText");
      await context.sync();

      if (selectedShapes.items.length === 0) {
        console.log("No shapes selected.");
        event.completed();
        return;
      }

      selectedShapes.items.forEach((shape, i) => {
        if (shape.textFrame && shape.textFrame.hasText) {
          console.log(`Clearing text in selected shape #${i}`);
          shape.textFrame.textRange.text = "";
        }
      });

      await context.sync();
      console.log("Text cleared from selected shapes.");
    });
  } catch (error) {
    console.error("Error in emptyTextBoxes:", error);
  } finally {
    event.completed();
  }
}

export async function emptyEntireSlide(event: Office.AddinCommands.Event) {
  console.log("emptyEntireSlide button clicked!");

  try {
    await PowerPoint.run(async (context) => {
      const selectedSlides = context.presentation.getSelectedSlides();
      selectedSlides.load("items");
      await context.sync();

      if (selectedSlides.items.length === 0) {
        console.log("No slide selected.");
        event.completed();
        return;
      }

      const slide = selectedSlides.items[0];
      const shapes = slide.shapes;
      shapes.load("items/textFrame/hasText");
      await context.sync();

      shapes.items.forEach((shape, i) => {
        if (shape.textFrame && shape.textFrame.hasText) {
          console.log(`Clearing shape #${i}`);
          shape.textFrame.textRange.text = "";
        }
      });

      await context.sync();
      console.log("Text cleared from shapes on selected slide.");
    });
  } catch (error) {
    console.error("Error in emptyEntireSlide:", error);
  } finally {
    event.completed();
  }
}

export async function equalizeSize(event: Office.AddinCommands.Event) {
  console.log("Equalize Size button clicked!");

  try {
    await PowerPoint.run(async (context) => {
      // Get the selected shapes on the slide
      const selectedShapes = context.presentation.getSelectedShapes();
      selectedShapes.load("items");
      await context.sync();

      if (selectedShapes.items.length === 0) {
        console.log("No shapes selected.");
        event.completed();
        return;
      }

      // Use the first shape as reference
      const referenceShape = selectedShapes.items[0];
      referenceShape.load(["width", "height"]);
      await context.sync();

      const refWidth = referenceShape.width;
      const refHeight = referenceShape.height;
      console.log(`Reference dimensions - Width: ${refWidth}, Height: ${refHeight}`);

      // Loop through the rest of the selected shapes and apply the same size
      for (let i = 1; i < selectedShapes.items.length; i++) {
        const shape = selectedShapes.items[i];
        shape.width = refWidth;
        shape.height = refHeight;
      }

      await context.sync();
      console.log("All selected shapes have been resized to match the first shape.");
    });
  } catch (error) {
    console.error("Error in equalizeSize:", error);
  } finally {
    event.completed();
  }
}

export async function equalizeHeight(event: Office.AddinCommands.Event) {
  console.log("Equalize Height button clicked!");

  try {
    await PowerPoint.run(async (context) => {
      // Get the selected shapes on the slide
      const selectedShapes = context.presentation.getSelectedShapes();
      selectedShapes.load("items");
      await context.sync();

      if (selectedShapes.items.length === 0) {
        console.log("No shapes selected.");
        event.completed();
        return;
      }

      // Use the first shape as reference
      const referenceShape = selectedShapes.items[0];
      referenceShape.load(["height"]);
      await context.sync();

      const refHeight = referenceShape.height;
      console.log(`Reference dimensions Height: ${refHeight}`);

      // Loop through the rest of the selected shapes and apply the same height
      for (let i = 1; i < selectedShapes.items.length; i++) {
        const shape = selectedShapes.items[i];
        shape.height = refHeight;
      }

      await context.sync();
      console.log("All selected shapes have been resized to match the first shape's height.");
    });
  } catch (error) {
    console.error("Error in equalizeHeight:", error);
  } finally {
    event.completed();
  }
}

export async function equalizeWidth(event: Office.AddinCommands.Event) {
  console.log("Equalize Size button clicked!");

  try {
    await PowerPoint.run(async (context) => {
      // Get the selected shapes on the slide
      const selectedShapes = context.presentation.getSelectedShapes();
      selectedShapes.load("items");
      await context.sync();

      if (selectedShapes.items.length === 0) {
        console.log("No shapes selected.");
        event.completed();
        return;
      }

      // Use the first shape as reference
      const referenceShape = selectedShapes.items[0];
      referenceShape.load(["width"]);
      await context.sync();

      const refWidth = referenceShape.width;
      console.log(`Reference dimensions - Width: ${refWidth}`);

      // Loop through the rest of the selected shapes and apply the same size
      for (let i = 1; i < selectedShapes.items.length; i++) {
        const shape = selectedShapes.items[i];
        shape.width = refWidth;
      }

      await context.sync();
      console.log("All selected shapes have been resized to match the first shape's width.");
    });
  } catch (error) {
    console.error("Error in equalizeSize:", error);
  } finally {
    event.completed();
  }
}

export async function stackHorizontally(event: Office.AddinCommands.Event) {
  console.log("Stack Horizontally button clicked!");

  try {
    await PowerPoint.run(async (context) => {
      // Retrieve the selected shapes
      const selectedShapes = context.presentation.getSelectedShapes();
      selectedShapes.load("items");
      await context.sync();

      if (selectedShapes.items.length === 0) {
        console.log("No shapes selected.");
        event.completed();
        return;
      }

      // Load required properties (left, width, and top) for all selected shapes
      selectedShapes.items.forEach((shape) => {
        shape.load(["left", "width", "top"]);
      });
      await context.sync();

      // Use the first shape's left coordinate as the starting point
      let currentLeft = selectedShapes.items[0].left;
      // Optionally, align all shapes vertically to the first shape's top coordinate
      const commonTop = selectedShapes.items[0].top;

      // Loop through each shape and reposition them in a row (end to end)
      for (let i = 0; i < selectedShapes.items.length; i++) {
        const shape = selectedShapes.items[i];
        shape.left = currentLeft;
        shape.top = commonTop;
        currentLeft += shape.width;
      }

      await context.sync();
      console.log("Selected shapes have been stacked horizontally.");
    });
  } catch (error) {
    console.error("Error in stackHorizontally:", error);
  } finally {
    event.completed();
  }
}

export async function stackVertically(event: Office.AddinCommands.Event) {
  console.log("Stack Vertically button clicked!");

  try {
    await PowerPoint.run(async (context) => {
      // Retrieve the selected shapes
      const selectedShapes = context.presentation.getSelectedShapes();
      selectedShapes.load("items");
      await context.sync();

      if (selectedShapes.items.length === 0) {
        console.log("No shapes selected.");
        event.completed();
        return;
      }

      // Load required properties (top, height, and left) for all selected shapes
      selectedShapes.items.forEach((shape) => {
        shape.load(["top", "height", "left"]);
      });
      await context.sync();

      // Use the first shape's top coordinate as the starting point
      let currentTop = selectedShapes.items[0].top;
      // Optionally, align all shapes horizontally to the first shape's left coordinate
      const commonLeft = selectedShapes.items[0].left;

      // Loop through each shape and reposition them in a column (stacked vertically)
      for (let i = 0; i < selectedShapes.items.length; i++) {
        const shape = selectedShapes.items[i];
        shape.top = currentTop;
        shape.left = commonLeft;
        currentTop += shape.height;
      }

      await context.sync();
      console.log("Selected shapes have been stacked vertically.");
    });
  } catch (error) {
    console.error("Error in stackVertically:", error);
  } finally {
    event.completed();
  }
}

export async function invertPositions(event: Office.AddinCommands.Event) {
  console.log("Invert Positions button clicked!");

  try {
    await PowerPoint.run(async (context) => {
      // Get the selected shapes
      const selectedShapes = context.presentation.getSelectedShapes();
      selectedShapes.load("items");
      await context.sync();

      if (selectedShapes.items.length !== 2) {
        console.log("Please select exactly two shapes.");
        event.completed();
        return;
      }

      // Retrieve the two shapes
      const shape1 = selectedShapes.items[0];
      const shape2 = selectedShapes.items[1];

      // Load the current left and top values for both shapes
      shape1.load(["left", "top"]);
      shape2.load(["left", "top"]);
      await context.sync();

      // Store shape1's current position
      const tempLeft = shape1.left;
      const tempTop = shape1.top;

      // Swap positions
      shape1.left = shape2.left;
      shape1.top = shape2.top;
      shape2.left = tempLeft;
      shape2.top = tempTop;

      await context.sync();
      console.log("The positions of the two shapes have been swapped.");
    });
  } catch (error) {
    console.error("Error in invertPositions:", error);
  } finally {
    event.completed();
  }
}

export async function removeMargins(event: Office.AddinCommands.Event) {
  console.log("Remove Margins button clicked!");

  try {
    await PowerPoint.run(async (context) => {
      // Get the selected shapes on the slide
      const selectedShapes = context.presentation.getSelectedShapes();
      selectedShapes.load("items");
      await context.sync();

      if (selectedShapes.items.length === 0) {
        console.log("No shapes selected.");
        event.completed();
        return;
      }

      // Loop through each shape that has a textFrame and set its margins to 0
      selectedShapes.items.forEach((shape) => {
        if (shape.textFrame) {
          // Assuming the textFrame object supports these margin properties.
          shape.textFrame.leftMargin = 0;
          shape.textFrame.rightMargin = 0;
          shape.textFrame.topMargin = 0;
          shape.textFrame.bottomMargin = 0;
        }
      });

      await context.sync();
      console.log("Margins removed from selected shapes.");
    });
  } catch (error) {
    console.error("Error in removeMargins:", error);
  } finally {
    event.completed();
  }
}

export async function toggleTextWrap(event: Office.AddinCommands.Event) {
  console.log("Toggle Text Wrap button clicked!");

  try {
    await PowerPoint.run(async (context) => {
      // Get the currently selected shapes on the slide
      const selectedShapes = context.presentation.getSelectedShapes();
      selectedShapes.load("items");
      await context.sync();

      if (selectedShapes.items.length === 0) {
        console.log("No shapes selected.");
        event.completed();
        return;
      }

      // Load the current wordWrap property for each shape that has a text frame.
      selectedShapes.items.forEach((shape) => {
        if (shape.textFrame) {
          shape.textFrame.load("wordWrap");
        }
      });
      await context.sync();

      // Toggle the wordWrap property for each shape with a text frame.
      selectedShapes.items.forEach((shape) => {
        if (shape.textFrame) {
          // Toggle the property: if currently true, set to false; if false, set to true.
          shape.textFrame.wordWrap = !shape.textFrame.wordWrap;
          console.log(`Text wrap set to: ${shape.textFrame.wordWrap}`);
        }
      });

      await context.sync();
      console.log("Toggled text wrap for selected shapes.");
    });
  } catch (error) {
    console.error("Error toggling text wrap:", error);
  } finally {
    event.completed();
  }
}

export async function copyPosition(event: Office.AddinCommands.Event) {
  console.log("Copy Position button clicked!");

  try {
    await PowerPoint.run(async (context) => {
      // Get the selected shapes on the slide.
      const selectedShapes = context.presentation.getSelectedShapes();
      selectedShapes.load("items");
      await context.sync();

      if (selectedShapes.items.length === 0) {
        console.log("No shapes selected for copying position.");
        event.completed();
        return;
      }

      // Use the first selected shape as reference.
      const shape = selectedShapes.items[0];
      shape.load(["left", "top"]);
      await context.sync();

      // Create a position object.
      const pos = {
        left: shape.left,
        top: shape.top,
      };

      // Persist the copied position using OfficeRuntime.storage.
      await OfficeRuntime.storage.setItem("copiedPosition", JSON.stringify(pos));
      console.log(`Position copied: left=${pos.left}, top=${pos.top}`);
    });
  } catch (error) {
    console.error("Error in copyPosition:", error);
  } finally {
    event.completed();
  }
}

export async function pastePosition(event: Office.AddinCommands.Event) {
  console.log("Paste Position button clicked!");

  try {
    // Retrieve the stored position using OfficeRuntime.storage.
    const posStr = await OfficeRuntime.storage.getItem("copiedPosition");
    if (!posStr) {
      console.log("No position has been copied yet.");
      event.completed();
      return;
    }
    const copiedPosition = JSON.parse(posStr);

    await PowerPoint.run(async (context) => {
      // Get the selected shapes on the slide.
      const selectedShapes = context.presentation.getSelectedShapes();
      selectedShapes.load("items");
      await context.sync();

      if (selectedShapes.items.length === 0) {
        console.log("No shapes selected for pasting position.");
        event.completed();
        return;
      }

      // Use the first selected shape and apply the stored position.
      const shape = selectedShapes.items[0];
      shape.left = copiedPosition.left;
      shape.top = copiedPosition.top;
      console.log(`Pasted position: left=${copiedPosition.left}, top=${copiedPosition.top}`);

      await context.sync();
    });
  } catch (error) {
    console.error("Error in pastePosition:", error);
  } finally {
    event.completed();
  }
}

export async function copyDimensions(event: Office.AddinCommands.Event) {
  console.log("Copy Dimensions button clicked!");

  try {
    await PowerPoint.run(async (context) => {
      // Get the currently selected shapes on the slide.
      const selectedShapes = context.presentation.getSelectedShapes();
      selectedShapes.load("items");
      await context.sync();

      if (selectedShapes.items.length === 0) {
        console.log("No shapes selected for copying dimensions.");
        event.completed();
        return;
      }

      // Use the first selected shape as reference.
      const shape = selectedShapes.items[0];
      shape.load(["width", "height"]);
      await context.sync();

      // Create an object with the dimensions.
      const dims = {
        width: shape.width,
        height: shape.height,
      };

      // Save the dimensions to OfficeRuntime.storage so that they persist.
      await OfficeRuntime.storage.setItem("copiedDimensions", JSON.stringify(dims));
      console.log(`Dimensions copied: width=${dims.width}, height=${dims.height}`);
    });
  } catch (error) {
    console.error("Error in copyDimensions:", error);
  } finally {
    event.completed();
  }
}

export async function pasteDimensions(event: Office.AddinCommands.Event) {
  console.log("Paste Dimensions button clicked!");

  try {
    // Retrieve the stored dimensions using OfficeRuntime.storage.
    const dimsStr = await OfficeRuntime.storage.getItem("copiedDimensions");
    if (!dimsStr) {
      console.log("No dimensions have been copied yet.");
      event.completed();
      return;
    }
    const copiedDimensions = JSON.parse(dimsStr);

    await PowerPoint.run(async (context) => {
      // Get the currently selected shapes on the slide.
      const selectedShapes = context.presentation.getSelectedShapes();
      selectedShapes.load("items");
      await context.sync();

      if (selectedShapes.items.length === 0) {
        console.log("No shapes selected for pasting dimensions.");
        event.completed();
        return;
      }

      // Use the first selected shape to apply the stored dimensions.
      const shape = selectedShapes.items[0];
      shape.width = copiedDimensions.width;
      shape.height = copiedDimensions.height;
      console.log(`Pasted dimensions: width=${copiedDimensions.width}, height=${copiedDimensions.height}`);

      await context.sync();
    });
  } catch (error) {
    console.error("Error in pasteDimensions:", error);
  } finally {
    event.completed();
  }
}

export async function autoFitShapeToText(event: Office.AddinCommands.Event) {
  console.log("Auto Fit Shape to Text button clicked!");
  try {
    await PowerPoint.run(async (context) => {
      // Get the currently selected shapes on the slide.
      const selectedShapes = context.presentation.getSelectedShapes();
      selectedShapes.load("items");
      await context.sync();

      if (selectedShapes.items.length === 0) {
        console.log("No shapes selected.");
        event.completed();
        return;
      }

      // Loop through each selected shape and set its textFrame.autoFit property.
      selectedShapes.items.forEach((shape) => {
        if (shape.textFrame) {
          // Resize the shape to exactly fit the text.
          shape.textFrame.autoSizeSetting = "AutoSizeShapeToFitText";
          console.log(`ShapeToFitText set for shape with id: ${shape.id}`);
        }
      });

      await context.sync();
      console.log("Auto fit shape to text completed.");
    });
  } catch (error) {
    console.error("Error in autoFitShapeToText:", error);
  } finally {
    event.completed();
  }
}

export async function autoFitTextToShape(event: Office.AddinCommands.Event) {
  console.log("Auto Fit Text to Shape button clicked!");
  try {
    await PowerPoint.run(async (context) => {
      // Get the currently selected shapes on the slide.
      const selectedShapes = context.presentation.getSelectedShapes();
      selectedShapes.load("items");
      await context.sync();

      if (selectedShapes.items.length === 0) {
        console.log("No shapes selected.");
        event.completed();
        return;
      }

      // Loop through each selected shape and set its textFrame.autoFit property.
      selectedShapes.items.forEach((shape) => {
        if (shape.textFrame) {
          // Adjust the text so that it fits within the shape (e.g. by shrinking text on overflow).
          shape.textFrame.autoSizeSetting = "AutoSizeTextToFitShape";
          console.log(`ShrinkTextOnOverflow set for shape with id: ${shape.id}`);
        }
      });

      await context.sync();
      console.log("Auto fit text to shape completed.");
    });
  } catch (error) {
    console.error("Error in autoFitTextToShape:", error);
  } finally {
    event.completed();
  }
}

export async function insertHarvey0(event: Office.AddinCommands.Event) {
  console.log("Insert Harvey Ball into selection button clicked!");
  const heavyCircle = "⭘"; // Harvey ball (Unicode U+2B58)

  try {
    await PowerPoint.run(async (context) => {
      // Retrieve the currently selected shapes.
      const selectedShapes = context.presentation.getSelectedShapes();
      selectedShapes.load("items");
      await context.sync();

      let inserted = false;

      if (selectedShapes.items.length > 0) {
        // First, try to update existing shapes that have a text frame.
        for (let i = 0; i < selectedShapes.items.length; i++) {
          const shape = selectedShapes.items[i];
          // Check if the shape has a text frame.
          if (shape.textFrame) {
            // Load the current text.
            shape.textFrame.load("textRange/text");
          }
        }
        await context.sync();

        // Append the heavy circle to each shape that has a text frame.
        for (let i = 0; i < selectedShapes.items.length; i++) {
          const shape = selectedShapes.items[i];
          if (shape.textFrame) {
            let currentText = shape.textFrame.textRange.text;
            // Append the heavy circle.
            shape.textFrame.textRange.text = currentText + heavyCircle;
            inserted = true;
          }
        }
        await context.sync();
      }

      // If no shape was updated (or nothing was selected), create a new text box.
      if (!inserted) {
        // Retrieve the currently selected slides.
        const selectedSlides = context.presentation.getSelectedSlides();
        selectedSlides.load("items");
        await context.sync();

        if (selectedSlides.items.length > 0) {
          // Use the first selected slide.
          const slide = selectedSlides.items[0];
          // Create a new text box with the heavy circle.
          const textShape = slide.shapes.addTextBox(heavyCircle);
          // Optionally set the position for the new text box.
          textShape.left = 100;
          textShape.top = 100;
          console.log("No suitable text frame found; created new text box with Harvey ball.");
        } else {
          console.log("No slide available to insert the text box.");
        }
      } else {
        console.log("Harvey ball inserted into existing text selection.");
      }
    });
  } catch (error) {
    console.error("Error inserting Harvey ball:", error);
  } finally {
    event.completed();
  }
}

export async function insertHarvey25(event: Office.AddinCommands.Event) {
  console.log("Insert Havey25 button clicked!");
  const symbol = "◔"; // Circle with upper right quadrant black (Unicode U+25D4)

  try {
    await PowerPoint.run(async (context) => {
      // Retrieve the currently selected shapes.
      const selectedShapes = context.presentation.getSelectedShapes();
      selectedShapes.load("items");
      await context.sync();

      let inserted = false;

      if (selectedShapes.items.length > 0) {
        // Attempt to insert the symbol into shapes that have a text frame.
        for (let i = 0; i < selectedShapes.items.length; i++) {
          const shape = selectedShapes.items[i];
          if (shape.textFrame) {
            // Load the current text.
            shape.textFrame.load("textRange/text");
          }
        }
        await context.sync();

        // Append the symbol to each shape with a text frame.
        for (let i = 0; i < selectedShapes.items.length; i++) {
          const shape = selectedShapes.items[i];
          if (shape.textFrame) {
            const currentText = shape.textFrame.textRange.text || "";
            shape.textFrame.textRange.text = currentText + symbol;
            inserted = true;
          }
        }
        await context.sync();
      }

      // If no shape was updated (or nothing was selected), create a new text box.
      if (!inserted) {
        // Retrieve the currently selected slides.
        const selectedSlides = context.presentation.getSelectedSlides();
        selectedSlides.load("items");
        await context.sync();

        if (selectedSlides.items.length > 0) {
          // Use the first selected slide.
          const slide = selectedSlides.items[0];
          // Create a new text box with the symbol.
          const textShape = slide.shapes.addTextBox(symbol);
          // Optionally, set the position for the new text box.
          textShape.left = 100;
          textShape.top = 100;
          console.log("No suitable text frame found; created new text box with symbol.");
        } else {
          console.log("No slide available to insert the text box.");
        }
      } else {
        console.log("Symbol inserted into existing text selection.");
      }
    });
  } catch (error) {
    console.error("Error inserting symbol:", error);
  } finally {
    event.completed();
  }
}

export async function insertHarvey50r(event: Office.AddinCommands.Event) {
  console.log("Insert Harvey50 button clicked!");
  const symbol = "◑"; // Circle with upper right quadrant black (Unicode U+25D4)

  try {
    await PowerPoint.run(async (context) => {
      // Retrieve the currently selected shapes.
      const selectedShapes = context.presentation.getSelectedShapes();
      selectedShapes.load("items");
      await context.sync();

      let inserted = false;

      if (selectedShapes.items.length > 0) {
        // Attempt to insert the symbol into shapes that have a text frame.
        for (let i = 0; i < selectedShapes.items.length; i++) {
          const shape = selectedShapes.items[i];
          if (shape.textFrame) {
            // Load the current text.
            shape.textFrame.load("textRange/text");
          }
        }
        await context.sync();

        // Append the symbol to each shape with a text frame.
        for (let i = 0; i < selectedShapes.items.length; i++) {
          const shape = selectedShapes.items[i];
          if (shape.textFrame) {
            const currentText = shape.textFrame.textRange.text || "";
            shape.textFrame.textRange.text = currentText + symbol;
            inserted = true;
          }
        }
        await context.sync();
      }

      // If no shape was updated (or nothing was selected), create a new text box.
      if (!inserted) {
        // Retrieve the currently selected slides.
        const selectedSlides = context.presentation.getSelectedSlides();
        selectedSlides.load("items");
        await context.sync();

        if (selectedSlides.items.length > 0) {
          // Use the first selected slide.
          const slide = selectedSlides.items[0];
          // Create a new text box with the symbol.
          const textShape = slide.shapes.addTextBox(symbol);
          // Optionally, set the position for the new text box.
          textShape.left = 100;
          textShape.top = 100;
          console.log("No suitable text frame found; created new text box with symbol.");
        } else {
          console.log("No slide available to insert the text box.");
        }
      } else {
        console.log("Symbol inserted into existing text selection.");
      }
    });
  } catch (error) {
    console.error("Error inserting symbol:", error);
  } finally {
    event.completed();
  }
}

export async function insertHarvey50l(event: Office.AddinCommands.Event) {
  console.log("Insert Harvey50 button clicked!");
  const symbol = "◐"; // Circle with upper right quadrant black (Unicode U+25D4)

  try {
    await PowerPoint.run(async (context) => {
      // Retrieve the currently selected shapes.
      const selectedShapes = context.presentation.getSelectedShapes();
      selectedShapes.load("items");
      await context.sync();

      let inserted = false;

      if (selectedShapes.items.length > 0) {
        // Attempt to insert the symbol into shapes that have a text frame.
        for (let i = 0; i < selectedShapes.items.length; i++) {
          const shape = selectedShapes.items[i];
          if (shape.textFrame) {
            // Load the current text.
            shape.textFrame.load("textRange/text");
          }
        }
        await context.sync();

        // Append the symbol to each shape with a text frame.
        for (let i = 0; i < selectedShapes.items.length; i++) {
          const shape = selectedShapes.items[i];
          if (shape.textFrame) {
            const currentText = shape.textFrame.textRange.text || "";
            shape.textFrame.textRange.text = currentText + symbol;
            inserted = true;
          }
        }
        await context.sync();
      }

      // If no shape was updated (or nothing was selected), create a new text box.
      if (!inserted) {
        // Retrieve the currently selected slides.
        const selectedSlides = context.presentation.getSelectedSlides();
        selectedSlides.load("items");
        await context.sync();

        if (selectedSlides.items.length > 0) {
          // Use the first selected slide.
          const slide = selectedSlides.items[0];
          // Create a new text box with the symbol.
          const textShape = slide.shapes.addTextBox(symbol);
          // Optionally, set the position for the new text box.
          textShape.left = 100;
          textShape.top = 100;
          console.log("No suitable text frame found; created new text box with symbol.");
        } else {
          console.log("No slide available to insert the text box.");
        }
      } else {
        console.log("Symbol inserted into existing text selection.");
      }
    });
  } catch (error) {
    console.error("Error inserting symbol:", error);
  } finally {
    event.completed();
  }
}

export async function insertHarvey75(event: Office.AddinCommands.Event) {
  console.log("Insert Harvey50 button clicked!");
  const symbol = "◕"; // Circle with upper right quadrant black (Unicode U+25D4)

  try {
    await PowerPoint.run(async (context) => {
      // Retrieve the currently selected shapes.
      const selectedShapes = context.presentation.getSelectedShapes();
      selectedShapes.load("items");
      await context.sync();

      let inserted = false;

      if (selectedShapes.items.length > 0) {
        // Attempt to insert the symbol into shapes that have a text frame.
        for (let i = 0; i < selectedShapes.items.length; i++) {
          const shape = selectedShapes.items[i];
          if (shape.textFrame) {
            // Load the current text.
            shape.textFrame.load("textRange/text");
          }
        }
        await context.sync();

        // Append the symbol to each shape with a text frame.
        for (let i = 0; i < selectedShapes.items.length; i++) {
          const shape = selectedShapes.items[i];
          if (shape.textFrame) {
            const currentText = shape.textFrame.textRange.text || "";
            shape.textFrame.textRange.text = currentText + symbol;
            inserted = true;
          }
        }
        await context.sync();
      }

      // If no shape was updated (or nothing was selected), create a new text box.
      if (!inserted) {
        // Retrieve the currently selected slides.
        const selectedSlides = context.presentation.getSelectedSlides();
        selectedSlides.load("items");
        await context.sync();

        if (selectedSlides.items.length > 0) {
          // Use the first selected slide.
          const slide = selectedSlides.items[0];
          // Create a new text box with the symbol.
          const textShape = slide.shapes.addTextBox(symbol);
          // Optionally, set the position for the new text box.
          textShape.left = 100;
          textShape.top = 100;
          console.log("No suitable text frame found; created new text box with symbol.");
        } else {
          console.log("No slide available to insert the text box.");
        }
      } else {
        console.log("Symbol inserted into existing text selection.");
      }
    });
  } catch (error) {
    console.error("Error inserting symbol:", error);
  } finally {
    event.completed();
  }
}

export async function insertHarvey100(event: Office.AddinCommands.Event) {
  console.log("Insert Harvey50 button clicked!");
  const symbol = "⬤"; // Circle with upper right quadrant black (Unicode U+25D4)

  try {
    await PowerPoint.run(async (context) => {
      // Retrieve the currently selected shapes.
      const selectedShapes = context.presentation.getSelectedShapes();
      selectedShapes.load("items");
      await context.sync();

      let inserted = false;

      if (selectedShapes.items.length > 0) {
        // Attempt to insert the symbol into shapes that have a text frame.
        for (let i = 0; i < selectedShapes.items.length; i++) {
          const shape = selectedShapes.items[i];
          if (shape.textFrame) {
            // Load the current text.
            shape.textFrame.load("textRange/text");
          }
        }
        await context.sync();

        // Append the symbol to each shape with a text frame.
        for (let i = 0; i < selectedShapes.items.length; i++) {
          const shape = selectedShapes.items[i];
          if (shape.textFrame) {
            const currentText = shape.textFrame.textRange.text || "";
            shape.textFrame.textRange.text = currentText + symbol;
            inserted = true;
          }
        }
        await context.sync();
      }

      // If no shape was updated (or nothing was selected), create a new text box.
      if (!inserted) {
        // Retrieve the currently selected slides.
        const selectedSlides = context.presentation.getSelectedSlides();
        selectedSlides.load("items");
        await context.sync();

        if (selectedSlides.items.length > 0) {
          // Use the first selected slide.
          const slide = selectedSlides.items[0];
          // Create a new text box with the symbol.
          const textShape = slide.shapes.addTextBox(symbol);
          // Optionally, set the position for the new text box.
          textShape.left = 100;
          textShape.top = 100;
          console.log("No suitable text frame found; created new text box with symbol.");
        } else {
          console.log("No slide available to insert the text box.");
        }
      } else {
        console.log("Symbol inserted into existing text selection.");
      }
    });
  } catch (error) {
    console.error("Error inserting symbol:", error);
  } finally {
    event.completed();
  }
}

export async function insertGreenLight(event: Office.AddinCommands.Event) {
  console.log("Insert Harvey50 button clicked!");
  const symbol = "🟢"; // Circle with upper right quadrant black (Unicode U+25D4)

  try {
    await PowerPoint.run(async (context) => {
      // Retrieve the currently selected shapes.
      const selectedShapes = context.presentation.getSelectedShapes();
      selectedShapes.load("items");
      await context.sync();

      let inserted = false;

      if (selectedShapes.items.length > 0) {
        // Attempt to insert the symbol into shapes that have a text frame.
        for (let i = 0; i < selectedShapes.items.length; i++) {
          const shape = selectedShapes.items[i];
          if (shape.textFrame) {
            // Load the current text.
            shape.textFrame.load("textRange/text");
          }
        }
        await context.sync();

        // Append the symbol to each shape with a text frame.
        for (let i = 0; i < selectedShapes.items.length; i++) {
          const shape = selectedShapes.items[i];
          if (shape.textFrame) {
            const currentText = shape.textFrame.textRange.text || "";
            shape.textFrame.textRange.text = currentText + symbol;
            inserted = true;
          }
        }
        await context.sync();
      }

      // If no shape was updated (or nothing was selected), create a new text box.
      if (!inserted) {
        // Retrieve the currently selected slides.
        const selectedSlides = context.presentation.getSelectedSlides();
        selectedSlides.load("items");
        await context.sync();

        if (selectedSlides.items.length > 0) {
          // Use the first selected slide.
          const slide = selectedSlides.items[0];
          // Create a new text box with the symbol.
          const textShape = slide.shapes.addTextBox(symbol);
          // Optionally, set the position for the new text box.
          textShape.left = 100;
          textShape.top = 100;
          console.log("No suitable text frame found; created new text box with symbol.");
        } else {
          console.log("No slide available to insert the text box.");
        }
      } else {
        console.log("Symbol inserted into existing text selection.");
      }
    });
  } catch (error) {
    console.error("Error inserting symbol:", error);
  } finally {
    event.completed();
  }
}

export async function insertOrangeLight(event: Office.AddinCommands.Event) {
  console.log("Insert Harvey50 button clicked!");
  const symbol = "🟠"; // Circle with upper right quadrant black (Unicode U+25D4)

  try {
    await PowerPoint.run(async (context) => {
      // Retrieve the currently selected shapes.
      const selectedShapes = context.presentation.getSelectedShapes();
      selectedShapes.load("items");
      await context.sync();

      let inserted = false;

      if (selectedShapes.items.length > 0) {
        // Attempt to insert the symbol into shapes that have a text frame.
        for (let i = 0; i < selectedShapes.items.length; i++) {
          const shape = selectedShapes.items[i];
          if (shape.textFrame) {
            // Load the current text.
            shape.textFrame.load("textRange/text");
          }
        }
        await context.sync();

        // Append the symbol to each shape with a text frame.
        for (let i = 0; i < selectedShapes.items.length; i++) {
          const shape = selectedShapes.items[i];
          if (shape.textFrame) {
            const currentText = shape.textFrame.textRange.text || "";
            shape.textFrame.textRange.text = currentText + symbol;
            inserted = true;
          }
        }
        await context.sync();
      }

      // If no shape was updated (or nothing was selected), create a new text box.
      if (!inserted) {
        // Retrieve the currently selected slides.
        const selectedSlides = context.presentation.getSelectedSlides();
        selectedSlides.load("items");
        await context.sync();

        if (selectedSlides.items.length > 0) {
          // Use the first selected slide.
          const slide = selectedSlides.items[0];
          // Create a new text box with the symbol.
          const textShape = slide.shapes.addTextBox(symbol);
          // Optionally, set the position for the new text box.
          textShape.left = 100;
          textShape.top = 100;
          console.log("No suitable text frame found; created new text box with symbol.");
        } else {
          console.log("No slide available to insert the text box.");
        }
      } else {
        console.log("Symbol inserted into existing text selection.");
      }
    });
  } catch (error) {
    console.error("Error inserting symbol:", error);
  } finally {
    event.completed();
  }
}

export async function insertRedLight(event: Office.AddinCommands.Event) {
  console.log("Insert Harvey50 button clicked!");
  const symbol = "🔴"; // Circle with upper right quadrant black (Unicode U+25D4)

  try {
    await PowerPoint.run(async (context) => {
      // Retrieve the currently selected shapes.
      const selectedShapes = context.presentation.getSelectedShapes();
      selectedShapes.load("items");
      await context.sync();

      let inserted = false;

      if (selectedShapes.items.length > 0) {
        // Attempt to insert the symbol into shapes that have a text frame.
        for (let i = 0; i < selectedShapes.items.length; i++) {
          const shape = selectedShapes.items[i];
          if (shape.textFrame) {
            // Load the current text.
            shape.textFrame.load("textRange/text");
          }
        }
        await context.sync();

        // Append the symbol to each shape with a text frame.
        for (let i = 0; i < selectedShapes.items.length; i++) {
          const shape = selectedShapes.items[i];
          if (shape.textFrame) {
            const currentText = shape.textFrame.textRange.text || "";
            shape.textFrame.textRange.text = currentText + symbol;
            inserted = true;
          }
        }
        await context.sync();
      }

      // If no shape was updated (or nothing was selected), create a new text box.
      if (!inserted) {
        // Retrieve the currently selected slides.
        const selectedSlides = context.presentation.getSelectedSlides();
        selectedSlides.load("items");
        await context.sync();

        if (selectedSlides.items.length > 0) {
          // Use the first selected slide.
          const slide = selectedSlides.items[0];
          // Create a new text box with the symbol.
          const textShape = slide.shapes.addTextBox(symbol);
          // Optionally, set the position for the new text box.
          textShape.left = 100;
          textShape.top = 100;
          console.log("No suitable text frame found; created new text box with symbol.");
        } else {
          console.log("No slide available to insert the text box.");
        }
      } else {
        console.log("Symbol inserted into existing text selection.");
      }
    });
  } catch (error) {
    console.error("Error inserting symbol:", error);
  } finally {
    event.completed();
  }
}

export async function insertUpArrow(event: Office.AddinCommands.Event) {
  console.log("Insert Harvey50 button clicked!");
  const symbol = "↑"; // Circle with upper right quadrant black (Unicode U+25D4)

  try {
    await PowerPoint.run(async (context) => {
      // Retrieve the currently selected shapes.
      const selectedShapes = context.presentation.getSelectedShapes();
      selectedShapes.load("items");
      await context.sync();

      let inserted = false;

      if (selectedShapes.items.length > 0) {
        // Attempt to insert the symbol into shapes that have a text frame.
        for (let i = 0; i < selectedShapes.items.length; i++) {
          const shape = selectedShapes.items[i];
          if (shape.textFrame) {
            // Load the current text.
            shape.textFrame.load("textRange/text");
          }
        }
        await context.sync();

        // Append the symbol to each shape with a text frame.
        for (let i = 0; i < selectedShapes.items.length; i++) {
          const shape = selectedShapes.items[i];
          if (shape.textFrame) {
            const currentText = shape.textFrame.textRange.text || "";
            shape.textFrame.textRange.text = currentText + symbol;
            inserted = true;
          }
        }
        await context.sync();
      }

      // If no shape was updated (or nothing was selected), create a new text box.
      if (!inserted) {
        // Retrieve the currently selected slides.
        const selectedSlides = context.presentation.getSelectedSlides();
        selectedSlides.load("items");
        await context.sync();

        if (selectedSlides.items.length > 0) {
          // Use the first selected slide.
          const slide = selectedSlides.items[0];
          // Create a new text box with the symbol.
          const textShape = slide.shapes.addTextBox(symbol);
          // Optionally, set the position for the new text box.
          textShape.left = 100;
          textShape.top = 100;
          console.log("No suitable text frame found; created new text box with symbol.");
        } else {
          console.log("No slide available to insert the text box.");
        }
      } else {
        console.log("Symbol inserted into existing text selection.");
      }
    });
  } catch (error) {
    console.error("Error inserting symbol:", error);
  } finally {
    event.completed();
  }
}

export async function insertDownArrow(event: Office.AddinCommands.Event) {
  console.log("Insert Harvey50 button clicked!");
  const symbol = "↓"; // Circle with upper right quadrant black (Unicode U+25D4)

  try {
    await PowerPoint.run(async (context) => {
      // Retrieve the currently selected shapes.
      const selectedShapes = context.presentation.getSelectedShapes();
      selectedShapes.load("items");
      await context.sync();

      let inserted = false;

      if (selectedShapes.items.length > 0) {
        // Attempt to insert the symbol into shapes that have a text frame.
        for (let i = 0; i < selectedShapes.items.length; i++) {
          const shape = selectedShapes.items[i];
          if (shape.textFrame) {
            // Load the current text.
            shape.textFrame.load("textRange/text");
          }
        }
        await context.sync();

        // Append the symbol to each shape with a text frame.
        for (let i = 0; i < selectedShapes.items.length; i++) {
          const shape = selectedShapes.items[i];
          if (shape.textFrame) {
            const currentText = shape.textFrame.textRange.text || "";
            shape.textFrame.textRange.text = currentText + symbol;
            inserted = true;
          }
        }
        await context.sync();
      }

      // If no shape was updated (or nothing was selected), create a new text box.
      if (!inserted) {
        // Retrieve the currently selected slides.
        const selectedSlides = context.presentation.getSelectedSlides();
        selectedSlides.load("items");
        await context.sync();

        if (selectedSlides.items.length > 0) {
          // Use the first selected slide.
          const slide = selectedSlides.items[0];
          // Create a new text box with the symbol.
          const textShape = slide.shapes.addTextBox(symbol);
          // Optionally, set the position for the new text box.
          textShape.left = 100;
          textShape.top = 100;
          console.log("No suitable text frame found; created new text box with symbol.");
        } else {
          console.log("No slide available to insert the text box.");
        }
      } else {
        console.log("Symbol inserted into existing text selection.");
      }
    });
  } catch (error) {
    console.error("Error inserting symbol:", error);
  } finally {
    event.completed();
  }
}

export async function insertEqualsSign(event: Office.AddinCommands.Event) {
  console.log("Insert Harvey50 button clicked!");
  const symbol = "＝"; // Circle with upper right quadrant black (Unicode U+25D4)

  try {
    await PowerPoint.run(async (context) => {
      // Retrieve the currently selected shapes.
      const selectedShapes = context.presentation.getSelectedShapes();
      selectedShapes.load("items");
      await context.sync();

      let inserted = false;

      if (selectedShapes.items.length > 0) {
        // Attempt to insert the symbol into shapes that have a text frame.
        for (let i = 0; i < selectedShapes.items.length; i++) {
          const shape = selectedShapes.items[i];
          if (shape.textFrame) {
            // Load the current text.
            shape.textFrame.load("textRange/text");
          }
        }
        await context.sync();

        // Append the symbol to each shape with a text frame.
        for (let i = 0; i < selectedShapes.items.length; i++) {
          const shape = selectedShapes.items[i];
          if (shape.textFrame) {
            const currentText = shape.textFrame.textRange.text || "";
            shape.textFrame.textRange.text = currentText + symbol;
            inserted = true;
          }
        }
        await context.sync();
      }

      // If no shape was updated (or nothing was selected), create a new text box.
      if (!inserted) {
        // Retrieve the currently selected slides.
        const selectedSlides = context.presentation.getSelectedSlides();
        selectedSlides.load("items");
        await context.sync();

        if (selectedSlides.items.length > 0) {
          // Use the first selected slide.
          const slide = selectedSlides.items[0];
          // Create a new text box with the symbol.
          const textShape = slide.shapes.addTextBox(symbol);
          // Optionally, set the position for the new text box.
          textShape.left = 100;
          textShape.top = 100;
          console.log("No suitable text frame found; created new text box with symbol.");
        } else {
          console.log("No slide available to insert the text box.");
        }
      } else {
        console.log("Symbol inserted into existing text selection.");
      }
    });
  } catch (error) {
    console.error("Error inserting symbol:", error);
  } finally {
    event.completed();
  }
}

export async function createPostItNote(event: Office.AddinCommands.Event) {
  try {
    await PowerPoint.run(async (context) => {
      // Retrieve the currently selected slides.
      const slides = context.presentation.getSelectedSlides();
      slides.load("items");
      await context.sync();

      if (slides.items.length === 0) {
        console.log("No slide selected. Please select a slide.");
        event.completed();
        return;
      }

      // Use the first selected slide.
      const slide = slides.items[0];

      // Assume default slide dimensions (in points). Here, slideWidth is 960 and slideHeight is 540.
      const slideWidth = 960;
      const slideHeight = 540;

      // Define the desired note dimensions.
      const noteWidth = 250;
      const noteHeight = 250;

      // Calculate positions to center the square horizontally and vertically.
      const noteLeft = (slideWidth - noteWidth) / 2;
      const noteTop = (slideHeight - noteHeight) / 2;

      // Create a geometric shape (a rectangle) for the post-it note.
      const postIt = slide.shapes.addGeometricShape("Rectangle");
      await context.sync();

      // Set the dimensions and position of the post-it note.
      postIt.width = noteWidth;
      postIt.height = noteHeight;
      postIt.left = noteLeft;
      postIt.top = noteTop;

      // Set the fill color to #E6BD01.
      postIt.fill.setSolidColor("#E6BD01");

      // Set the text inside the shape to be black.
      postIt.textFrame.textRange.font.color = "black";

      // Optionally, hide the border.
      postIt.lineFormat.visible = false;

      await context.sync();
      console.log("Post-it note square created and centered on the slide.");
    });
  } catch (error) {
    console.error("Error creating post-it note square:", error);
  } finally {
    event.completed();
  }
}

export async function insertFootnote(event: Office.AddinCommands.Event) {
  try {
    await PowerPoint.run(async (context) => {
      // Retrieve the currently selected slides.
      const slides = context.presentation.getSelectedSlides();
      slides.load("items");
      await context.sync();

      if (slides.items.length === 0) {
        console.log("No slide selected. Please select a slide.");
        event.completed();
        return;
      }
      
      // Use the first selected slide.
      const slide = slides.items[0];

      // Load the presentation's slide dimensions.
      context.presentation.load(["slideWidth", "slideHeight"]);
      await context.sync();

      // Use actual slide dimensions if available, otherwise default values.
      const slideWidth = context.presentation.slideWidth || 960;
      const slideHeight = context.presentation.slideHeight || 540;

      // Define the desired text box dimensions.
      const boxWidth = 300;
      const boxHeight = 30;
      // Set margins: 20 points from the left, 20 points from the bottom.
      const left = 20;
      const top = slideHeight - boxHeight - 20;

      // Create a text box with the default footnote text.
      const footnoteShape = slide.shapes.addTextBox("Footnote text");
      await context.sync();

      // Set its size and position.
      footnoteShape.width = boxWidth;
      footnoteShape.height = boxHeight;
      footnoteShape.left = left;
      footnoteShape.top = top;

      // Format the text: set font size to 10 and italicize it.
      footnoteShape.textFrame.textRange.font.size = 10;
      footnoteShape.textFrame.textRange.font.italic = true;
      // Optionally, set the text color to black.
      footnoteShape.textFrame.textRange.font.color = "black";

      await context.sync();
      console.log("Footnote inserted on the bottom left of the slide.");
    });
  } catch (error) {
    console.error("Error inserting footnote:", error);
  } finally {
    event.completed();
  }
}

export async function createPreliminaryStamp(event: Office.AddinCommands.Event) {
  try {
    await PowerPoint.run(async (context) => {
      // Retrieve the currently selected slides.
      const slides = context.presentation.getSelectedSlides();
      slides.load("items");
      await context.sync();

      if (slides.items.length === 0) {
        console.log("No slide selected. Please select a slide.");
        event.completed();
        return;
      }
      // Use the first selected slide.
      const slide = slides.items[0];

      // Load the presentation's slide dimensions.
      context.presentation.load(["slideWidth", "slideHeight"]);
      await context.sync();
      const slideWidth = context.presentation.slideWidth || 960;
      const slideHeight = context.presentation.slideHeight || 540;

      // Define the stamp dimensions.
      const stampWidth = 120;
      const stampHeight = 50;
      const margin = 20;

      // Calculate the stamp's position in the top right corner.
      const stampLeft = slideWidth - stampWidth - margin;
      const stampTop = margin;

      // Create a geometric shape (rectangle) for the stamp.
      const stamp = slide.shapes.addGeometricShape("Rectangle");
      await context.sync();

      // Set the stamp's dimensions and position.
      stamp.width = stampWidth;
      stamp.height = stampHeight;
      stamp.left = stampLeft;
      stamp.top = stampTop;

      // Set the fill color to black.
      stamp.fill.setSolidColor("black");

      // Insert the text "PRELIMINARY" into the stamp.
      stamp.textFrame.textRange.text = "PRELIMINARY";
      // Format the text: white color, bold, and a font size of 16.
      stamp.textFrame.textRange.font.color = "white";
      stamp.textFrame.textRange.font.bold = true;
      stamp.textFrame.textRange.font.size = 16;
      stamp.textFrame.textRange.paragraphFormat.horizontalAlignment = "Center";
      stamp.lineFormat.visible = false;

      
      // Set autoFit so the shape resizes to fit the text.
      stamp.textFrame.autoSizeSetting = "AutoSizeShapeToFitText";

      await context.sync();
      console.log("Preliminary stamp created in the top right corner of the slide.");
    });
  } catch (error) {
    console.error("Error creating preliminary stamp:", error);
  } finally {
    event.completed();
  }
}

export async function createExampleStamp(event: Office.AddinCommands.Event) {
  try {
    await PowerPoint.run(async (context) => {
      // Retrieve the currently selected slides.
      const slides = context.presentation.getSelectedSlides();
      slides.load("items");
      await context.sync();

      if (slides.items.length === 0) {
        console.log("No slide selected. Please select a slide.");
        event.completed();
        return;
      }
      // Use the first selected slide.
      const slide = slides.items[0];

      // Load the presentation's slide dimensions.
      context.presentation.load(["slideWidth", "slideHeight"]);
      await context.sync();
      const slideWidth = context.presentation.slideWidth || 960;
      const slideHeight = context.presentation.slideHeight || 540;

      // Define the stamp dimensions.
      const stampWidth = 120;
      const stampHeight = 50;
      const margin = 20;

      // Calculate the stamp's position in the top right corner.
      const stampLeft = slideWidth - stampWidth - margin;
      const stampTop = margin;

      // Create a geometric shape (rectangle) for the stamp.
      const stamp = slide.shapes.addGeometricShape("Rectangle");
      await context.sync();

      // Set the stamp's dimensions and position.
      stamp.width = stampWidth;
      stamp.height = stampHeight;
      stamp.left = stampLeft;
      stamp.top = stampTop;

      // Set the fill color to black.
      stamp.fill.setSolidColor("green");

      // Insert the text "PRELIMINARY" into the stamp.
      stamp.textFrame.textRange.text = "EXAMPLE";
      // Format the text: white color, bold, and a font size of 16.
      stamp.textFrame.textRange.font.color = "white";
      stamp.textFrame.textRange.font.bold = true;
      stamp.textFrame.textRange.font.size = 16;
      stamp.textFrame.textRange.paragraphFormat.horizontalAlignment = "Center";
      stamp.lineFormat.visible = false;

      
      // Set autoFit so the shape resizes to fit the text.
      stamp.textFrame.autoSizeSetting = "AutoSizeShapeToFitText";

      await context.sync();
      console.log("Preliminary stamp created in the top right corner of the slide.");
    });
  } catch (error) {
    console.error("Error creating preliminary stamp:", error);
  } finally {
    event.completed();
  }
}

export async function createExecSumStamp(event: Office.AddinCommands.Event) {
  try {
    await PowerPoint.run(async (context) => {
      // Retrieve the currently selected slides.
      const slides = context.presentation.getSelectedSlides();
      slides.load("items");
      await context.sync();

      if (slides.items.length === 0) {
        console.log("No slide selected. Please select a slide.");
        event.completed();
        return;
      }
      // Use the first selected slide.
      const slide = slides.items[0];

      // Load the presentation's slide dimensions.
      context.presentation.load(["slideWidth", "slideHeight"]);
      await context.sync();
      const slideWidth = context.presentation.slideWidth || 960;
      const slideHeight = context.presentation.slideHeight || 540;

      // Define the stamp dimensions.
      const stampWidth = 120;
      const stampHeight = 50;
      const margin = 20;

      // Calculate the stamp's position in the top right corner.
      const stampLeft = slideWidth - stampWidth - margin;
      const stampTop = margin;

      // Create a geometric shape (rectangle) for the stamp.
      const stamp = slide.shapes.addGeometricShape("Rectangle");
      await context.sync();

      // Set the stamp's dimensions and position.
      stamp.width = stampWidth;
      stamp.height = stampHeight;
      stamp.left = stampLeft;
      stamp.top = stampTop;

      // Set the fill color to black.
      stamp.fill.setSolidColor("red");

      // Insert the text "PRELIMINARY" into the stamp.
      stamp.textFrame.textRange.text = "EXECUTIVE SUMMARY";
      // Format the text: white color, bold, and a font size of 16.
      stamp.textFrame.textRange.font.color = "white";
      stamp.textFrame.textRange.font.bold = true;
      stamp.textFrame.textRange.font.size = 16;
      stamp.textFrame.textRange.paragraphFormat.horizontalAlignment = "Center";
      stamp.lineFormat.visible = false;

      
      // Set autoFit so the shape resizes to fit the text.
      stamp.textFrame.autoSizeSetting = "AutoSizeShapeToFitText";

      await context.sync();
      console.log("Preliminary stamp created in the top right corner of the slide.");
    });
  } catch (error) {
    console.error("Error creating preliminary stamp:", error);
  } finally {
    event.completed();
  }
}

export async function createDraftStamp(event: Office.AddinCommands.Event) {
  try {
    await PowerPoint.run(async (context) => {
      // Retrieve the currently selected slides.
      const slides = context.presentation.getSelectedSlides();
      slides.load("items");
      await context.sync();

      if (slides.items.length === 0) {
        console.log("No slide selected. Please select a slide.");
        event.completed();
        return;
      }
      // Use the first selected slide.
      const slide = slides.items[0];

      // Load the presentation's slide dimensions.
      context.presentation.load(["slideWidth", "slideHeight"]);
      await context.sync();
      const slideWidth = context.presentation.slideWidth || 960;
      const slideHeight = context.presentation.slideHeight || 540;

      // Define the stamp dimensions.
      const stampWidth = 80;
      const stampHeight = 50;
      const margin = 20;

      // Calculate the stamp's position in the top right corner.
      const stampLeft = slideWidth - stampWidth - margin;
      const stampTop = margin;

      // Create a geometric shape (rectangle) for the stamp.
      const stamp = slide.shapes.addGeometricShape("Rectangle");
      await context.sync();

      // Set the stamp's dimensions and position.
      stamp.width = stampWidth;
      stamp.height = stampHeight;
      stamp.left = stampLeft;
      stamp.top = stampTop;

      // Set the fill color to black.
      stamp.fill.setSolidColor("CCA801");

      // Insert the text "PRELIMINARY" into the stamp.
      stamp.textFrame.textRange.text = "DRAFT";
      // Format the text: white color, bold, and a font size of 16.
      stamp.textFrame.textRange.font.color = "white";
      stamp.textFrame.textRange.font.bold = true;
      stamp.textFrame.textRange.font.size = 16;
      stamp.textFrame.textRange.paragraphFormat.horizontalAlignment = "Center";
      stamp.lineFormat.visible = false;

      
      // Set autoFit so the shape resizes to fit the text.
      stamp.textFrame.autoSizeSetting = "AutoSizeShapeToFitText";

      await context.sync();
      console.log("Preliminary stamp created in the top right corner of the slide.");
    });
  } catch (error) {
    console.error("Error creating preliminary stamp:", error);
  } finally {
    event.completed();
  }
}


