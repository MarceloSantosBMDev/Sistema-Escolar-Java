package escola;

public class Aluno {
    private String nome;
    private float nota;
    private boolean fezProva;

    public Aluno(String nome, boolean fezProva, float nota) {
        this.nome = nome;
        this.fezProva =fezProva;
        this.nota = nota;
    }

    public String getNome() { 
        return nome; 
    }
    
    public float getNota() { 
        return nota; 
    }
    
    public boolean getFezProva() { 
        return fezProva; 
    }

    @Override
    public String toString() {
        return nome + " - " + (fezProva ? "Nota: " + nota : "Não fez prova");
    }
}