import pandas as pd
df = pd.DataFrame({'fact': ['Нефть образовалась из древних организмов', 'Первая нефтяная скважина - 1859 год', 'Россия - лидер по добыче']})
df.to_excel('facts.xlsx', index=False)
print("facts.xlsx created")
