
/*=====================================================
        Retrieve the file array buffer
=======================================================*/
/**
 * Get the local file as an array buffer.
 *@param fileElementId String that specifies the element ID.
*/
export function getFileBuffer(fileElementId: string) {
    var fileInput:any = $('#' + fileElementId);
    var deferred = $.Deferred();
    var reader = new FileReader();
    reader.onloadend = function (e:any) {
    deferred.resolve(e.target.result);
    }
    reader.onerror = function (e:any) {
    deferred.reject(e.target.error);
    }
    reader.readAsArrayBuffer(fileInput[0].files[0]);
    return deferred.promise();
}