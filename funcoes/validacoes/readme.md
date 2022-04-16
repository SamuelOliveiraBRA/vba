### **VALIDAÇÕES DE CPF OU CNPJ**

## Instalação
Para realizar a instalação, basta criar um arquivo com extensão **.bas** e depois importar para o seu projeto através do Menu do aplicativo do office na **Janela de Propriedades** e clicar em *Arquivo >  Importar Arquivo* ou pressionar *CTRL + M*

<div>
  <img alt="Propriedade do Sistema" src="https://doutorexcel.files.wordpress.com/2011/03/editor-vba1.jpg"/>
</div>

#### validarCPF()
-Realiza a valiação do CPF de acordo com o número que é passado como parâmetro
> o valor de retorno original é o valor de **False** e medida que é validado, ele retorna como **True**

Para utlizar, basta chamar através de uma string ou função, neste exemplo de string com o exemplo **minhaString = _validarCPF( número do CPF )_**

- _o número do CPF_ pode ser passado como '000000000' ou '000.000.000-00'.


#### validarCNPJ()
-Realiza a valiação do CNPJ de acordo com o número que é passado como parâmetro
> o valor de retorno original é o valor de **False** e medida que é validado, ele retorna como **True**

Para utlizar, basta chamar através de uma string ou função, neste exemplo de string com o exemplo **minhaString = _validarCNPJ( número do CNPJ )_**

- _o número do CNPJ_ pode ser passado como '00000000000000' ou '00.000.000/0000-00'.
