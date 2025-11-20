import pandas as pd
import json

def excel_to_json(excel_file='Curated Chaos.xlsx', output_file='products.json'):
    """
    Convert Excel product inventory to JSON format for the website
    """
    try:
        # Read the Excel file
        df = pd.read_excel(excel_file)
        
        # Print column names for debugging
        print("📋 Found columns:", list(df.columns))
        print(f"📊 Found {len(df)} rows\n")
        
        # Initialize products list
        products = []
        
        # Process each row
        for idx, row in df.iterrows():
            # Check if name exists and is not empty
            if pd.isna(row.get('name')) or str(row.get('name')).strip() == '':
                print(f"⚠️  Skipping row {idx+2} - no name provided")
                continue
            
            # Collect image filenames (image1, image2, image3)
            images = []
            for col in ['image1', 'image2', 'image3']:
                if col in df.columns and pd.notna(row.get(col)) and str(row.get(col)).strip() != '':
                    images.append(str(row[col]).strip())
            
            # Skip if no images
            if not images:
                print(f"⚠️  Skipping '{row.get('name')}' - no images provided")
                continue
            
            # Get category, default to 'misc' if not provided
            category = str(row.get('category', 'misc')).strip().lower()
            if category not in ['women', 'men', 'books', 'misc']:
                print(f"⚠️  '{row.get('name')}' has invalid category '{category}', defaulting to 'misc'")
                category = 'misc'
            
            # Build product object
            product = {
                'name': str(row['name']).strip(),
                'category': category,
                'images': images,
                'price': int(float(row.get('price', 0))),
                'condition': str(row.get('condition', 'Good')).strip()
            }
            
            # Add optional fields if they exist and are not empty
            if 'size' in df.columns and pd.notna(row.get('size')) and str(row.get('size')).strip() != '':
                product['size'] = str(row['size']).strip()
            
            if 'brand' in df.columns and pd.notna(row.get('brand')) and str(row.get('brand')).strip() != '':
                product['brand'] = str(row['brand']).strip()
            
            if 'author' in df.columns and pd.notna(row.get('author')) and str(row.get('author')).strip() != '':
                product['author'] = str(row['author']).strip()
            
            if 'description' in df.columns and pd.notna(row.get('description')) and str(row.get('description')).strip() != '':
                product['description'] = str(row['description']).strip()
            
            products.append(product)
            print(f"✅ Added: {product['name']} ({len(images)} image(s))")
        
        # Write to JSON file
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(products, f, indent=2, ensure_ascii=False)
        
        print(f"\n🎉 Success! Converted {len(products)} products to {output_file}")
        
        if len(products) > 0:
            print(f"\nProducts by category:")
            # Show summary
            categories = {}
            for p in products:
                cat = p['category']
                categories[cat] = categories.get(cat, 0) + 1
            
            for cat, count in categories.items():
                print(f"  - {cat}: {count} items")
        else:
            print("\n⚠️  No products were converted. Please check your Excel file.")
        
        return True
        
    except FileNotFoundError:
        print(f"❌ Error: Could not find '{excel_file}'")
        print("Make sure your Excel file is in the same folder as this script!")
        return False
    
    except Exception as e:
        print(f"❌ Error: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("🔄 Converting Excel to JSON...\n")
    excel_to_json()
    print("\n✨ Done! Your products.json is ready to use!")