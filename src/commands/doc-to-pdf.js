/* global Office, Word */

function getSliceAsync(file, nextSlice, sliceCount, gotAllSlices, docdataSlices, slicesReceived) {
  file.getSliceAsync(nextSlice, function (sliceResult) {
    if (sliceResult.status == Office.AsyncResultStatus.Succeeded) {
      if (!gotAllSlices) {
        // Failed to get all slices, no need to continue.
        return;
      }

      // Got one slice, store it in a temporary array.
      // (Or you can do something else, such as
      // send it to a third-party server.)
      docdataSlices[sliceResult.value.index] = sliceResult.value.data;
      if (++slicesReceived === sliceCount) {
        // All slices have been received.
        file.closeAsync();
        onGotAllSlices(docdataSlices);
      } else {
        getSliceAsync(file, ++nextSlice, sliceCount, gotAllSlices, docdataSlices, slicesReceived);
      }
    } else {
      gotAllSlices = false;
      file.closeAsync();
      Word.showNotification("getSliceAsync Error:", sliceResult.error.message);
    }
  });
}

function onGotAllSlices(docdataSlices) {
  let docdata = [];
  for (let i = 0; i < docdataSlices.length; i++) {
    docdata = docdata.concat(docdataSlices[i]);
  }

  let fileContent = String();
  for (let j = 0; j < docdata.length; j++) {
    fileContent += String.fromCharCode(docdata[j]);
  }

  // Now all the file content is stored in 'fileContent' variable,
  // you can do something with it, such as print, fax...

  //Print file content
  Word.run(async (context) => {
    const paragraph = context.document.body.insertParagraph(fileContent, Word.InsertLocation.end);

    paragraph.font.color = "red";

    await context.sync();
  });
}

function getDocumentAsPdf() {
  // The following example gets the document in PDF format.
  Office.context.document.getFileAsync(Office.FileType.Pdf, { sliceSize: 65536 }, function (result) {
    if (result.status == Office.AsyncResultStatus.Succeeded) {
      Word.run(async (context) => {
        const paragraph = context.document.body.insertParagraph("step 1 success", Word.InsertLocation.end);

        paragraph.font.color = "red";

        await context.sync();
      });
      const file = result.value;
      const sliceCount = file.sliceCount;
      const slicesReceived = 0,
        gotAllSlices = true,
        docdataSlices = [];
      Word.showNotification("File size:" + file.size + " #Slices: " + sliceCount);
      // Now, you can call getSliceAsync to download the files,
      // as described in the previous code segment (compressed format).
      getSliceAsync(file, 0, sliceCount, gotAllSlices, docdataSlices, slicesReceived);
      file.closeAsync();
    } else {
      Word.run(async (context) => {
        const paragraph = context.document.body.insertParagraph(JSON.stringify(result), Word.InsertLocation.end);

        paragraph.font.color = "red";

        await context.sync();
      });
      Word.showNotification("Error:", result.error.message);
    }
  });
}

export default getDocumentAsPdf;
