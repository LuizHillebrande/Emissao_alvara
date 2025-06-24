# Sistema de Salvamento de Caminhos de Pastas

## Como Funciona

O sistema agora salva automaticamente os caminhos das pastas que você digita na interface, para que na próxima vez que abrir o aplicativo, os caminhos ainda estejam lá.

### Arquivos Criados

O sistema cria dois arquivos de texto para salvar os caminhos:

- `caminho_maringa.txt` - Salva o caminho da pasta para Maringá
- `caminho_tapejara.txt` - Salva o caminho da pasta para Tapejara

### Funcionalidades

1. **Salvamento Automático**: Quando você inicia o processo (clica no botão), o caminho digitado é salvo automaticamente
2. **Carregamento Automático**: Quando você abre o aplicativo, os caminhos salvos são carregados automaticamente nos campos
3. **Flexibilidade**: Você pode modificar os caminhos a qualquer momento e eles serão salvos na próxima execução

### Exemplo de Uso

1. Digite um caminho como: `Z:\SOCIETÁRIO\TAXAS DE ALVARÁ\2025\MARINGA`
2. Clique em "Iniciar prefeitura de Maringá/PR"
3. O caminho é salvo automaticamente
4. Na próxima vez que abrir o app, o caminho ainda estará lá
5. Você pode modificar o caminho se quiser usar uma pasta diferente

### Formato dos Caminhos

Use o formato Windows padrão:
- `Z:\PASTA\SUBPASTA`
- `C:\Users\SeuUsuario\Desktop\Boletos`
- `\\servidor\compartilhamento\pasta`

### Arquivos de Progresso

O sistema também mantém os arquivos de progresso existentes:
- `progresso_maringa.txt` - Última linha processada de Maringá
- `progresso_tapejara.txt` - Última linha processada de Tapejara 