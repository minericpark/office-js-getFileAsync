/* global Office, Word */

function getSliceAsync(file, nextSlice, sliceCount, gotAllSlices, docdataSlices, slicesReceived) {
  file.getSliceAsync(nextSlice, function (sliceResult) {
    if (sliceResult.status == Office.AsyncResultStatus.Succeeded) {
      if (!gotAllSlices) {
        // Failed to get all slices, no need to continue.
        Word.run(async (context) => {
          const paragraph = context.document.body.insertParagraph("Failed to get slices", Word.InsertLocation.end);

          paragraph.font.color = "red";

          await context.sync();
        });
        return;
      }

      // Got one slice, store it in a temporary array.
      // (Or you can do something else, such as
      // send it to a third-party server.)
      docdataSlices[sliceResult.value.index] = sliceResult.value.data;
      if (++slicesReceived === sliceCount) {
        // All slices have been received.
        Word.run(async (context) => {
          const paragraph = context.document.body.insertParagraph("Slices get!", Word.InsertLocation.end);

          paragraph.font.color = "red";

          await context.sync();
        });
        file.closeAsync((result) => {
          Word.run(async (context) => {
            const paragraph = context.document.body.insertParagraph(
              "close in getSlice: " + JSON.stringify(result),
              Word.InsertLocation.end
            );

            paragraph.font.color = "red";

            await context.sync();
          });
        });
        onGotAllSlices(docdataSlices);
      } else {
        getSliceAsync(file, ++nextSlice, sliceCount, gotAllSlices, docdataSlices, slicesReceived);
      }
    } else {
      gotAllSlices = false;
      file.closeAsync((result) => {
        Word.run(async (context) => {
          const paragraph = context.document.body.insertParagraph(
            "close in failed getSlice: " + JSON.stringify(result),
            Word.InsertLocation.end
          );

          paragraph.font.color = "red";

          await context.sync();
        });
      });
      Word.run(async (context) => {
        const paragraph = context.document.body.insertParagraph("Error in getSliceAsync", Word.InsertLocation.end);

        paragraph.font.color = "red";

        await context.sync();
      });
    }
  });
}

function onGotAllSlices(docdataSlices) {
  let docdata = [];
  for (let i = 0; i < docdataSlices.length; i++) {
    docdata = docdata.concat(docdataSlices[i]);
  }

  let fileContent = String();
  // Now all the file content is stored in 'fileContent' variable
  for (let j = 0; j < docdata.length; j++) {
    fileContent += String.fromCharCode(docdata[j]);
  }

  printFileContent(docdata);
}

function printFileContent(fileContent) {
  let formData = new FormData();
  let blob = new Blob([new Uint8Array(fileContent)], { type: "application/pdf" });
  formData.append("file", blob);
  Word.run(async (context) => {
    let object = {};
    formData.forEach((value, key) => {
      // Reflect.has in favor of: object.hasOwnProperty(key)
      if (!Reflect.has(object, key)) {
        object[key] = value;
        return;
      }
      if (!Array.isArray(object[key])) {
        object[key] = [object[key]];
      }
      object[key].push(value);
    });
    const json = JSON.stringify(object);
    const paragraph = context.document.body.insertParagraph("fileContent: " + json, Word.InsertLocation.end);

    paragraph.font.color = "red";

    await context.sync();
  });
}

function getDocumentAsPdf() {
  Office.context.document.getFileAsync(Office.FileType.Pdf, { sliceSize: 65536 }, function (result) {
    if (result.status == Office.AsyncResultStatus.Succeeded) {
      const file = result.value;
      const sliceCount = file.sliceCount;
      const slicesReceived = 0,
        gotAllSlices = true,
        docdataSlices = [];
      //Print file details
      Word.run(async (context) => {
        const paragraph = context.document.body.insertParagraph(
          "file: " + JSON.stringify(file),
          Word.InsertLocation.end
        );

        paragraph.font.color = "red";

        await context.sync();
      });
      //Assemble file slices into file object
      getSliceAsync(file, 0, sliceCount, gotAllSlices, docdataSlices, slicesReceived);
      //Removed closeAsync; led to more consistency, but seems necessary?
      file.closeAsync((result) => {
        Word.run(async (context) => {
          const paragraph = context.document.body.insertParagraph(
            "close in failed getDoc: " + JSON.stringify(result),
            Word.InsertLocation.end
          );

          paragraph.font.color = "red";

          await context.sync();
        });
      });
    } else {
      Word.run(async (context) => {
        const paragraph = context.document.body.insertParagraph(JSON.stringify(result), Word.InsertLocation.end);

        paragraph.font.color = "red";

        await context.sync();
      });
    }
  });
}

export default getDocumentAsPdf;
