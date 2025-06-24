package escola;

import java.util.ArrayList;
import java.util.List;

public class Sala {
    private int ano;
    private List<Aluno> alunos;

    public Sala(int ano) {
        this.ano = ano;
        this.alunos = new ArrayList<>();
    }

    public void adicionarAluno(Aluno aluno) {
        alunos.add(aluno);
    }

    public List<Aluno> getAlunos() {
        return alunos;
    }

    public int getAno() {
        return ano;
    }

    @Override
    public String toString() {
        return "Sala " + ano + "º Ano (" + alunos.size() + " alunos)";
    }
}