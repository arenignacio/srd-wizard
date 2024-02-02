import imageSize from "image-size";

/**
 * Get size as percentage of original size
 * @param {*} doc 
 */
export default function resize (image, targetWidth, targetHeight) {
   
    const originalSize = imageSize(image);
    const originalAspectRatio = originalSize.width / originalSize.height;

    let width;
    let height;

    if (!targetWidth) {
        // fixed height, calc width
        height = targetHeight;
        width = height * originalAspectRatio;
    } else if (!targetHeight) {
        // fixed width, calc height
        width = targetWidth;
        height = width / originalAspectRatio;
    } else {
        const targetRatio = targetWidth / targetHeight;
        if (targetRatio > originalAspectRatio) {
            // fill height, calc width
            height = targetHeight;
            width = height * originalAspectRatio;
        } else {
            // fill width, calc height
            width = targetWidth;
            height = width / originalAspectRatio;
        }
    }

    console.log(originalSize, originalAspectRatio, width, height);

    return Media.addImage(doc, image, width, height);
}