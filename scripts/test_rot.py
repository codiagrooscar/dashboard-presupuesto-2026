import json

with open('web_dashboard/production_planning_data.json', encoding='utf-8') as f:
    d = json.load(f)

base = d['base_skus']
nomovers = [b for b in base if b.get('is_nomover')]
slowmovers = [b for b in base if b.get('is_slowmover')]
fastmovers = [b for b in base if b.get('is_fastmover')]
total_stock_u = sum(b.get('stock_actual', 0) for b in base)
total_nomover_u = sum(b.get('no_mover_qty', 0) for b in base)
total_slowmover_u = sum(b.get('slow_mover_qty', 0) for b in base)
total_fastmover_u = sum(b.get('fast_mover_qty', 0) for b in base)

val_nomover = sum(b.get('no_mover_eur', 0) for b in base)
val_slowmover = sum(b.get('slow_mover_eur', 0) for b in base)
val_stock = sum(b.get('stock_eur', 0) for b in base)

print(f"Total Stock Físico Envasado: {total_stock_u:,.0f} u (Valor: {val_stock:,.2f} EUR)")
print(f"NoMovers (>6 meses nacional): {len(nomovers)} SKUs | {total_nomover_u:,.0f} u ({total_nomover_u/total_stock_u*100:.1f}%) | {val_nomover:,.2f} EUR")
print(f"SlowMovers (3 a 6 meses nac): {len(slowmovers)} SKUs | {total_slowmover_u:,.0f} u ({total_slowmover_u/total_stock_u*100:.1f}%) | {val_slowmover:,.2f} EUR")
print(f"Alta Rotacion (0 a 3 m nac): {len(fastmovers)} SKUs | {total_fastmover_u:,.0f} u ({total_fastmover_u/total_stock_u*100:.1f}%)")
