"""
CLI entry point

Converts draw.io files to PowerPoint presentations
"""
import sys
import argparse
from pathlib import Path

from drawio2pptx.io.drawio_loader import DrawIOLoader
from drawio2pptx.io.pptx_writer import PPTXWriter
from drawio2pptx.logger import ConversionLogger
from drawio2pptx.analysis import compare_conversion
from drawio2pptx.config import ConversionConfig, default_config
from drawio2pptx.media.image_utils import (
    clear_image_cache,
    get_image_cache_stats,
    reset_image_cache_stats,
)


def main():
    """Main function"""
    parser = argparse.ArgumentParser(
        description='Convert draw.io files to PowerPoint presentations',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  drawio2pptx input.drawio output.pptx
  drawio2pptx input.drawio output.pptx --analyze
  drawio2pptx input.drawio output.pptx --no-cache
  drawio2pptx input.drawio output.pptx --clear-cache
  drawio2pptx --clear-cache
  drawio2pptx input.drawio output.pptx -a
        """
    )
    parser.add_argument('input', nargs='?', type=str, help='Path to input draw.io file')
    parser.add_argument('output', nargs='?', type=str, help='Path to output PowerPoint file')
    parser.add_argument('-a', '--analyze', action='store_true',
                       help='Display analysis results after conversion')
    parser.set_defaults(image_cache=True)
    parser.add_argument(
        '--cache',
        dest='image_cache',
        action='store_true',
        help='Enable local image cache for URL image fetch and SVG->PNG results (default: enabled)',
    )
    parser.add_argument(
        '--no-cache',
        dest='image_cache',
        action='store_false',
        help='Disable local image cache',
    )
    parser.add_argument(
        '--cache-dir',
        dest='image_cache_dir',
        type=str,
        default=None,
        help='Directory for local image cache (default: ~/.cache/drawio2pptx/images)',
    )
    parser.add_argument(
        '--clear-cache',
        dest='clear_image_cache',
        action='store_true',
        help='Clear local image cache before conversion starts',
    )
    
    args = parser.parse_args()
    
    if args.clear_image_cache and not args.input and not args.output:
        config = ConversionConfig()
        if args.image_cache_dir:
            config.image_cache_dir = args.image_cache_dir
        clear_image_cache(config.image_cache_dir)
        print(f"Cleared image cache: {config.image_cache_dir}")
        return

    if not args.input or not args.output:
        parser.error("the following arguments are required: input, output")

    input_path = Path(args.input)
    output_path = Path(args.output)
    
    if not input_path.exists():
        print(f"Input file not found: {input_path}")
        sys.exit(1)
    
    print(f"Parsing: {input_path}")
    
    try:
        # Create conversion configuration (default: 192 DPI, auto-calculated for minimum 100px short edge)
        config = ConversionConfig()
        config.image_cache_enabled = bool(args.image_cache)
        if args.image_cache_dir:
            config.image_cache_dir = args.image_cache_dir
        default_config.image_cache_enabled = config.image_cache_enabled
        default_config.image_cache_dir = config.image_cache_dir
        if args.clear_image_cache:
            clear_image_cache(config.image_cache_dir)
            print(f"Cleared image cache: {config.image_cache_dir}")
        
        # Create logger with config
        logger = ConversionLogger(config=config)
        if config.image_cache_enabled:
            logger.info(f"Image cache: enabled ({config.image_cache_dir})")
        else:
            logger.info("Image cache: disabled")
        reset_image_cache_stats()
        
        # Load draw.io file
        loader = DrawIOLoader(logger=logger, config=config)
        diagrams = loader.load_file(input_path)
        
        if not diagrams:
            print("No diagrams found in file")
            sys.exit(1)
        
        # Get page size (from first diagram)
        page_size = loader.extract_page_size(diagrams[0])
        
        # Create PowerPoint presentation
        writer = PPTXWriter(logger=logger, config=config)
        prs, blank_layout = writer.create_presentation(page_size)
        
        # Process each diagram
        slide_count = 0
        for mgm in diagrams:
            # Extract elements
            elements = loader.extract_elements(mgm)
            
            # Add to slide
            writer.add_slide(prs, blank_layout, elements)
            slide_count += 1
        
        # Save
        prs.save(output_path)
        print(f"Saved {output_path} ({slide_count} slides)")
        if config.image_cache_enabled:
            stats = get_image_cache_stats()
            logger.info(
                "Image cache stats: "
                f"hits={stats.get('hits', 0)}, "
                f"misses={stats.get('misses', 0)}, "
                f"writes={stats.get('writes', 0)}"
            )

        # Display warnings
        warnings = logger.get_warnings()
        if warnings:
            print(f"\nWarnings ({len(warnings)}):")
            for warning in warnings:
                print(f"  - {warning.message}")
        
        # Execute analysis if analysis option is specified
        if args.analyze:
            compare_conversion(input_path, output_path)
    
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
