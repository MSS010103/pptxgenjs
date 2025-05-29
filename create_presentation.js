const PptxGenJS = require("pptxgenjs");
const fs = require("fs");
const path = require("path");
const axios = require("axios");
const ytdl = require("ytdl-core");
const instagramGetUrl = require("instagram-url-direct");

class MediaPresentationGenerator {
  constructor() {
    this.pptx = new PptxGenJS();
    this.maxMediaPerSlide = 6;
    this.supportedImageFormats = [
      ".jpg",
      ".jpeg",
      ".png",
      ".gif",
      ".bmp",
      ".webp",
    ];
    this.supportedVideoFormats = [
      ".mp4",
      ".avi",
      ".mov",
      ".wmv",
      ".flv",
      ".webm",
    ];
  }

  // Initialize presentation with basic settings
  initializePresentation() {
    this.pptx.defineLayout({ name: "CUSTOM", width: 10, height: 7.5 });
    this.pptx.layout = "CUSTOM";
    console.log("Presentation initialized");
  }

  // Download image from Instagram
  async downloadInstagramMedia(url, filename) {
    try {
      const mediaUrl = await instagramGetUrl(url);
      const response = await axios({
        method: "GET",
        url: mediaUrl,
        responseType: "stream",
      });

      const filePath = path.join(__dirname, "downloads", filename);
      const writer = fs.createWriteStream(filePath);
      response.data.pipe(writer);

      return new Promise((resolve, reject) => {
        writer.on("finish", () => resolve(filePath));
        writer.on("error", reject);
      });
    } catch (error) {
      console.error("Error downloading Instagram media:", error);
      throw error;
    }
  }

  // Download video from YouTube
  async downloadYouTubeVideo(url, filename) {
    try {
      const filePath = path.join(__dirname, "downloads", filename);
      const stream = ytdl(url, {
        quality: "highest",
        filter: (format) => format.container === "mp4",
      });

      const writer = fs.createWriteStream(filePath);
      stream.pipe(writer);

      return new Promise((resolve, reject) => {
        writer.on("finish", () => resolve(filePath));
        writer.on("error", reject);
        stream.on("error", reject);
      });
    } catch (error) {
      console.error("Error downloading YouTube video:", error);
      throw error;
    }
  }

  // Convert file to base64
  convertToBase64(filePath) {
    try {
      const fileBuffer = fs.readFileSync(filePath);
      const fileExtension = path.extname(filePath).toLowerCase();
      let mimeType;

      // Determine MIME type
      switch (fileExtension) {
        case ".jpg":
        case ".jpeg":
          mimeType = "image/jpeg";
          break;
        case ".png":
          mimeType = "image/png";
          break;
        case ".gif":
          mimeType = "image/gif";
          break;
        case ".bmp":
          mimeType = "image/bmp";
          break;
        case ".webp":
          mimeType = "image/webp";
          break;
        case ".mp4":
          mimeType = "video/mp4";
          break;
        case ".avi":
          mimeType = "video/avi";
          break;
        case ".mov":
          mimeType = "video/mov";
          break;
        case ".wmv":
          mimeType = "video/wmv";
          break;
        case ".flv":
          mimeType = "video/flv";
          break;
        case ".webm":
          mimeType = "video/webm";
          break;
        default:
          mimeType = "application/octet-stream";
      }

      const base64String = `data:${mimeType};base64,${fileBuffer.toString(
        "base64"
      )}`;
      console.log(
        `Converted ${filePath} to base64 (${fileBuffer.length} bytes)`
      );
      return base64String;
    } catch (error) {
      console.error("Error converting file to base64:", error);
      throw error;
    }
  }

  // Get image dimensions from base64 data
  getImageDimensions(base64Data) {
    try {
      // Extract base64 string without data URL prefix
      const base64String = base64Data.split(",")[1];
      const buffer = Buffer.from(base64String, "base64");

      // Simple PNG dimension extraction
      if (base64Data.includes("image/png")) {
        const width = buffer.readUInt32BE(16);
        const height = buffer.readUInt32BE(20);
        return { width, height };
      }

      // Simple JPEG dimension extraction (basic approach)
      if (
        base64Data.includes("image/jpeg") ||
        base64Data.includes("image/jpg")
      ) {
        // For JPEG, we'll use a default aspect ratio approach
        // In production, you might want to use a proper image library
        return { width: 1920, height: 1080 }; // Default 16:9 ratio
      }

      // Default dimensions for other formats
      return { width: 1920, height: 1080 };
    } catch (error) {
      console.log("Could not extract dimensions, using defaults");
      return { width: 1920, height: 1080 };
    }
  }

  // Calculate optimal positioning for media items preserving aspect ratios
  calculateMediaPositions(mediaFiles) {
    const positions = [];
    const slideWidth = 10; // inches
    const slideHeight = 7.5; // inches
    const titleHeight = 1; // Space reserved for title
    const margin = 0.3;
    const spacing = 0.2; // Space between items

    const availableWidth = slideWidth - 2 * margin;
    const availableHeight = slideHeight - titleHeight - 2 * margin;
    const mediaCount = mediaFiles.length;

    let cols, rows;

    // Determine grid layout based on media count
    switch (mediaCount) {
      case 1:
        cols = 1;
        rows = 1;
        break;
      case 2:
        cols = 2;
        rows = 1;
        break;
      case 3:
        cols = 3;
        rows = 1;
        break;
      case 4:
        cols = 2;
        rows = 2;
        break;
      case 5:
        cols = 3;
        rows = 2;
        break;
      case 6:
        cols = 3;
        rows = 2;
        break;
      default:
        cols = 3;
        rows = 2;
    }

    // Calculate maximum available space per item
    const maxItemWidth = (availableWidth - spacing * (cols - 1)) / cols;
    const maxItemHeight = (availableHeight - spacing * (rows - 1)) / rows;

    for (let i = 0; i < mediaCount; i++) {
      const row = Math.floor(i / cols);
      const col = i % cols;

      // Get original dimensions if possible
      let originalWidth, originalHeight, aspectRatio;

      try {
        const base64Data = this.convertToBase64(mediaFiles[i].path);
        const dimensions = this.getImageDimensions(base64Data);
        originalWidth = dimensions.width;
        originalHeight = dimensions.height;
        aspectRatio = originalWidth / originalHeight;
      } catch (error) {
        // Default aspect ratio if we can't determine original
        aspectRatio = 16 / 9;
        originalWidth = 1920;
        originalHeight = 1080;
      }

      // Calculate size maintaining aspect ratio
      let itemWidth = maxItemWidth;
      let itemHeight = maxItemWidth / aspectRatio;

      // If height exceeds max, scale by height instead
      if (itemHeight > maxItemHeight) {
        itemHeight = maxItemHeight;
        itemWidth = maxItemHeight * aspectRatio;
      }

      // Calculate position to center the item in its grid cell
      const cellX = margin + col * (maxItemWidth + spacing);
      const cellY = titleHeight + margin + row * (maxItemHeight + spacing);

      const x = cellX + (maxItemWidth - itemWidth) / 2;
      const y = cellY + (maxItemHeight - itemHeight) / 2;

      positions.push({
        x: x,
        y: y,
        w: itemWidth,
        h: itemHeight,
        originalWidth: originalWidth,
        originalHeight: originalHeight,
        aspectRatio: aspectRatio,
      });
    }

    return positions;
  }

  // Add media to slide with proper positioning and preserved resolution
  addMediaToSlide(slide, mediaFiles) {
    const mediaCount = Math.min(mediaFiles.length, this.maxMediaPerSlide);
    const positions = this.calculateMediaPositions(
      mediaFiles.slice(0, mediaCount)
    );

    for (let i = 0; i < mediaCount; i++) {
      const mediaFile = mediaFiles[i];
      const position = positions[i];
      const fileExtension = path.extname(mediaFile.path).toLowerCase();

      try {
        const base64Data = this.convertToBase64(mediaFile.path);

        // Add image/video with preserved aspect ratio and no sizing constraints
        slide.addImage({
          data: base64Data,
          x: position.x,
          y: position.y,
          w: position.w,
          h: position.h,
          // Remove sizing property to prevent any scaling/fitting that might affect quality
          rounding: false, // Ensure no rounding that might affect pixels
        });

        console.log(`Added ${mediaFile.type} to slide:`);
        console.log(
          `  Position: (${position.x.toFixed(2)}, ${position.y.toFixed(2)})`
        );
        console.log(
          `  Size: ${position.w.toFixed(2)} x ${position.h.toFixed(2)} inches`
        );
        console.log(
          `  Original aspect ratio: ${position.aspectRatio.toFixed(2)}:1`
        );
      } catch (error) {
        console.error(`Error adding ${mediaFile.path} to slide:`, error);
      }
    }
  }

  // Process local media files
  async processLocalMedia(mediaDirectory) {
    const mediaFiles = [];

    try {
      const files = fs.readdirSync(mediaDirectory);

      for (const file of files) {
        const filePath = path.join(mediaDirectory, file);
        const fileExtension = path.extname(file).toLowerCase();
        const stats = fs.statSync(filePath);

        if (stats.isFile()) {
          let mediaType = "unknown";

          if (this.supportedImageFormats.includes(fileExtension)) {
            mediaType = "image";
          } else if (this.supportedVideoFormats.includes(fileExtension)) {
            mediaType = "video";
          }

          if (mediaType !== "unknown") {
            mediaFiles.push({
              path: filePath,
              type: mediaType,
              extension: fileExtension,
              name: file,
            });
          }
        }
      }

      console.log(`Found ${mediaFiles.length} media files`);
      return mediaFiles;
    } catch (error) {
      console.error("Error processing local media:", error);
      throw error;
    }
  }

  // Download media from URLs
  async downloadMediaFromUrls(urls) {
    const downloadDir = path.join(__dirname, "downloads");

    // Create downloads directory if it doesn't exist
    if (!fs.existsSync(downloadDir)) {
      fs.mkdirSync(downloadDir, { recursive: true });
    }

    const downloadedFiles = [];

    for (let i = 0; i < urls.length; i++) {
      const url = urls[i];
      try {
        let filePath;

        if (url.includes("instagram.com")) {
          const filename = `instagram_media_${i + 1}.jpg`;
          filePath = await this.downloadInstagramMedia(url, filename);
        } else if (url.includes("youtube.com") || url.includes("youtu.be")) {
          const filename = `youtube_video_${i + 1}.mp4`;
          filePath = await this.downloadYouTubeVideo(url, filename);
        } else {
          // Generic download for other URLs
          const filename = `media_${i + 1}${path.extname(url) || ".jpg"}`;
          const response = await axios({
            method: "GET",
            url: url,
            responseType: "stream",
          });

          filePath = path.join(downloadDir, filename);
          const writer = fs.createWriteStream(filePath);
          response.data.pipe(writer);

          await new Promise((resolve, reject) => {
            writer.on("finish", resolve);
            writer.on("error", reject);
          });
        }

        const fileExtension = path.extname(filePath).toLowerCase();
        let mediaType = "unknown";

        if (this.supportedImageFormats.includes(fileExtension)) {
          mediaType = "image";
        } else if (this.supportedVideoFormats.includes(fileExtension)) {
          mediaType = "video";
        }

        downloadedFiles.push({
          path: filePath,
          type: mediaType,
          extension: fileExtension,
          name: path.basename(filePath),
        });

        console.log(`Downloaded: ${url} -> ${filePath}`);
      } catch (error) {
        console.error(`Error downloading ${url}:`, error);
      }
    }

    return downloadedFiles;
  }

  // Create slides with media
  createSlidesWithMedia(mediaFiles) {
    const totalSlides = Math.ceil(mediaFiles.length / this.maxMediaPerSlide);

    for (let slideIndex = 0; slideIndex < totalSlides; slideIndex++) {
      const slide = this.pptx.addSlide();
      const startIndex = slideIndex * this.maxMediaPerSlide;
      const endIndex = Math.min(
        startIndex + this.maxMediaPerSlide,
        mediaFiles.length
      );
      const slideMediaFiles = mediaFiles.slice(startIndex, endIndex);

      // Add title to slide
      slide.addText(`Media Collection - Slide ${slideIndex + 1}`, {
        x: 0.5,
        y: 0.2,
        w: 9,
        h: 0.5,
        fontSize: 24,
        bold: true,
        align: "center",
      });

      // Add media to slide
      this.addMediaToSlide(slide, slideMediaFiles);

      console.log(
        `Created slide ${slideIndex + 1} with ${
          slideMediaFiles.length
        } media items`
      );
    }
  }

  // Generate and save presentation
  async generatePresentation(outputPath = "media_presentation.pptx") {
    try {
      await this.pptx.writeFile({ fileName: outputPath });
      console.log(`Presentation saved as: ${outputPath}`);
    } catch (error) {
      console.error("Error saving presentation:", error);
      throw error;
    }
  }

  // Main execution method
  async execute(options = {}) {
    const {
      urls = [],
      localMediaDirectory = null,
      outputFileName = "media_presentation.pptx",
    } = options;

    try {
      console.log("Starting media presentation generation...");

      // Initialize presentation
      this.initializePresentation();

      let allMediaFiles = [];

      // Download media from URLs if provided
      if (urls.length > 0) {
        console.log("Downloading media from URLs...");
        const downloadedFiles = await this.downloadMediaFromUrls(urls);
        allMediaFiles = allMediaFiles.concat(downloadedFiles);
      }

      // Process local media if directory provided
      if (localMediaDirectory) {
        console.log("Processing local media files...");
        const localFiles = await this.processLocalMedia(localMediaDirectory);
        allMediaFiles = allMediaFiles.concat(localFiles);
      }

      if (allMediaFiles.length === 0) {
        console.log("No media files found. Creating empty presentation.");
        const slide = this.pptx.addSlide();
        slide.addText("No Media Files Found", {
          x: 1,
          y: 3,
          w: 8,
          h: 1.5,
          fontSize: 32,
          align: "center",
        });
      } else {
        console.log(
          `Creating slides for ${allMediaFiles.length} media files...`
        );
        this.createSlidesWithMedia(allMediaFiles);
      }

      // Generate presentation
      await this.generatePresentation(outputFileName);

      console.log("Media presentation generation completed successfully!");
      return {
        success: true,
        totalMediaFiles: allMediaFiles.length,
        totalSlides: Math.ceil(allMediaFiles.length / this.maxMediaPerSlide),
        outputFile: outputFileName,
      };
    } catch (error) {
      console.error("Error in presentation generation:", error);
      throw error;
    }
  }
}

// Usage Example
async function main() {
  const generator = new MediaPresentationGenerator();

  try {
    // Option 1: Using local files only (recommended for testing)
    const localMediaDir = "./media_files"; // This folder should contain your images/videos

    const result = await generator.execute({
      urls: [], // Empty array for now
      localMediaDirectory: localMediaDir,
      outputFileName: "my_media_presentation.pptx",
    });

    console.log("Generation Result:", result);

    // Option 2: Using URLs (uncomment if needed)
    /*
        const urls = [
            'https://www.instagram.com/p/example1/',
            'https://www.youtube.com/watch?v=example1',
        ];
        
        const result = await generator.execute({
            urls: urls,
            localMediaDirectory: null,
            outputFileName: 'url_media_presentation.pptx'
        });
        */
  } catch (error) {
    console.error("Failed to generate presentation:", error);
  }
}

// Export the class for use in other files
module.exports = MediaPresentationGenerator;

// Run the example (uncomment to execute)
main();
