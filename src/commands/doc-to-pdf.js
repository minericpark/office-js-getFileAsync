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

          paragraph.font.color = "green";

          await context.sync();
        });
        file.closeAsync((result) => {
          Word.run(async (context) => {
            const paragraph = context.document.body.insertParagraph(
              "close in getSlice: " + JSON.stringify(result),
              Word.InsertLocation.end
            );

            paragraph.font.color = "green";

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

  /** Can store file content with charCode (string), kind of not necessary for us
  let fileContent = String();
  for (let j = 0; j < docdata.length; j++) {
    fileContent += String.fromCharCode(docdata[j]);
  }
   */
  printFileContent(docdata);
}

function printFileContent(fileContent) {
  let formData = new FormData();
  let blob = new Blob([new Uint8Array(fileContent)], { type: "application/pdf" });
  formData.append("file", blob);
  Word.run(async (context) => {
    //How to read formdata
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
    const paragraph = context.document.body.insertParagraph("fileContent: " + blob, Word.InsertLocation.end);

    //How to read base64 file data
    let reader = new FileReader();
    reader.readAsDataURL(blob);
    reader.onloadend = function () {
      let base64data = reader.result.toString();
      context.document.body.insertParagraph("base64: " + base64data.indexOf(",") + 1, Word.InsertLocation.end);
      //Only works with Compressed (Docx) base64 code, does not work with PDF
      context.document.body.insertFileFromBase64(
        base64data.substr(base64data.indexOf(",") + 1),
        Word.InsertLocation.end
      );
    };

    paragraph.font.color = "green";

    await context.sync();
  });
}

function getDocumentAsPdf() {
  //Either use Compressed to get Docx data slices, or PDF to get PDF data slices
  Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 65536 }, function (result) {
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

        paragraph.font.color = "green";

        await context.sync();
      });
      //Assemble file slices into file object
      getSliceAsync(file, 0, sliceCount, gotAllSlices, docdataSlices, slicesReceived);
      //Removed closeAsync; led to more consistency, but seems necessary?
      file.closeAsync((result) => {
        Word.run(async (context) => {
          const paragraph = context.document.body.insertParagraph(
            "close in getDoc: " + JSON.stringify(result),
            Word.InsertLocation.end
          );

          paragraph.font.color = "green";

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
