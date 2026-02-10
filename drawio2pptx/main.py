"""
CLI entry point

Converts draw.io files to PowerPoint presentations
"""
import sys
import argparse
from pathlib import Path

from drawio2pptx.io.drawio_loader import DrawIOLoader
from drawio2pptx.io.pptx_writer import PPTXWriter
from drawio2pptx.logger import get_logger, ConversionLogger
from drawio2pptx.analysis import compare_conversion
from drawio2pptx.config import ConversionConfig, default_config


def main():
    """Main function"""
    parser = argparse.ArgumentParser(
        description='Convert draw.io files to PowerPoint presentations',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  drawio2pptx input.drawio output.pptx
  drawio2pptx input.drawio output.pptx --analyze
  drawio2pptx input.drawio output.pptx -a
        """
    )
    parser.add_argument('input', type=str, help='Path to input draw.io file')
    parser.add_argument('output', type=str, help='Path to output PowerPoint file')
    parser.add_argument('-a', '--analyze', action='store_true',
                       help='Display analysis results after conversion')
    
    args = parser.parse_args()
    
    input_path = Path(args.input)
    output_path = Path(args.output)
    
    if not input_path.exists():
        print(f"Input file not found: {input_path}")
        sys.exit(1)
    
    print(f"Parsing: {input_path}")
    
    try:
        # Create conversion configuration
        config = ConversionConfig()  # Use default settings
        
        # Create logger with config
        logger = ConversionLogger(config=config)
        
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
