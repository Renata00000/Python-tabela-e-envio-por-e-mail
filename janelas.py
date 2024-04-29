from tkinter.filedialog import askdirectory , askopenfilenames
from tkinter.messagebox import askyesnocancel


# past_computador = askopenfilenames(title= "selecione uma pasta de computador")
# print(past_computador)

confirmacao = askyesnocancel(title='confirmacao', message= 'voce realmente quer iniciar essa autmacao')
print(confirmacao)
