# Guía de Contribución

## Antes de comenzar

1. Fork el repositorio
2. Clona tu fork localmente
3. Crea una rama para tu feature

## Proceso de desarrollo

1. **Instala las dependencias**:
   ```bash
   pip install -r requirements.txt
   ```

2. **Haz tus cambios**:
   - Mantén la coherencia con el estilo de código existente
   - No modifiques la lógica del negocio sin acuerdo
   - Actualiza la documentación si es necesario

3. **Prueba tus cambios**:
   ```bash
   streamlit run app.py
   # Prueba manualmente todos los flujos afectados
   ```

4. **Commit y Push**:
   ```bash
   git add .
   git commit -m "Descripción clara del cambio"
   git push origin feature/tu-feature
   ```

5. **Abre un Pull Request**:
   - Describe qué cambios hiciste y por qué
   - Referencia cualquier issue relacionado
   - Espera la revisión

## Lineamientos de código

- Usa nombres descriptivos para variables y funciones
- Añade docstrings a funciones públicas
- Mantén las líneas < 88 caracteres
- Usa type hints cuando sea posible

## Reporte de bugs

Si encuentras un bug:
1. Verifica que no esté ya reportado
2. Abre un issue con:
   - Descripción clara del problema
   - Pasos para reproducir
   - Resultado esperado vs actual
   - Tu entorno (Python version, OS, etc.)

## Sugerencias de mejoras

Las sugerencias son bienvenidas. Abre un issue con la etiqueta "enhancement" describiendo tu idea.

---

¡Gracias por contribuir! 🙌
