Public Class Form1

    Dim A(20, 20)
    Dim B(20)
    Dim C(20)
    Dim X(20)
    Dim R(20)
    Dim D(20, 20)
    Dim SUMX(20)
    Dim N = 3           'Numero di equazioni
    Dim NN = N - 1
    Dim K
    Dim V
    Dim I, J

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Me.BackColor = Color.Orange
        Me.Text = "Inversione di matrici metodo SIMEQ"
        Dim NuovoCarattere As New Font("Verdana", 12)
        TextBox1.Font = NuovoCarattere

        ' Matrice A(NxN) dei coefficienti
        A(1, 1) = 2
        A(1, 2) = 1
        A(1, 3) = 0
        A(2, 1) = 0
        A(2, 2) = 2
        A(2, 3) = 1
        A(3, 1) = 1
        A(3, 2) = 1
        A(3, 3) = 1

        ' Vettore B(N) delle costanti
        B(1) = 4
        B(2) = 2
        B(3) = 1

        ' Stampa A(NxN)
        TextBox1.Text &= ("Stampa matrice A(NxN) dei coefficienti per righe") & vbCrLf
        TextBox1.Text &= vbCrLf
        For I = 1 To N
            For J = 1 To N
                TextBox1.Text &= (A(I, J) & " ")
            Next J
            TextBox1.Text &= vbCrLf
        Next I
        TextBox1.Text &= vbCrLf

        ' Stampa B(N)
        TextBox1.Text &= ("Stampa vettore B(N) per righe") & vbCrLf
        TextBox1.Text &= vbCrLf
        For I = 1 To N
            TextBox1.Text &= (B(I) & vbCrLf)
        Next I
        TextBox1.Text &= vbCrLf

        Call SIMEQ()

        ' Stampa X(N)
        TextBox1.Text &= ("Stampa vettore soluzione X(N) per righe metodo SIMEQ") & vbCrLf
        TextBox1.Text &= vbCrLf
        For I = 1 To N
            TextBox1.Text &= (X(I) & vbCrLf)
        Next I
        TextBox1.Text &= vbCrLf

        Call MATIN()

        ' Stampa A(NxN)
        TextBox1.Text &= ("Stampa matrice inversa metodo MATIN A(NxN) per righe") & vbCrLf
        TextBox1.Text &= vbCrLf
        For I = 1 To N
            For J = 1 To N
                TextBox1.Text &= (A(I, J) & " ")
            Next J
            TextBox1.Text &= vbCrLf
        Next I
        TextBox1.Text &= vbCrLf

        ' Stampa X(N)
        TextBox1.Text &= ("Stampa vettore soluzione X(N) per righe metodo MATIN") & vbCrLf
        TextBox1.Text &= vbCrLf
        For I = 1 To N
            TextBox1.Text &= (X(I) & vbCrLf)
        Next I
        TextBox1.Text &= vbCrLf
    End Sub

    Private Sub SIMEQ()
        'Inizio iterazioni
        Dim Max = 50        'Numero max di iterazioni
        Dim ERR = 0.00001   'Convergenza
        Dim ITER = 1
        Dim SUM, LAST, INITL, TEMP
        Dim BIG
S53:    BIG = 0
        For I = 1 To N
            SUM = 0
            ' Questa sezione somma i termini di una riga escludendo quelli della diagonale principale
            If I = 1 Then GoTo S63
            LAST = I - 1
            For J = 1 To LAST
                SUM = SUM + A(I, J) * X(J)
            Next J
            If I = N Then GoTo S68
S63:        INITL = I + 1
            For J = INITL To N
                SUM = SUM + A(I, J) * X(J)
            Next J
            ' Calcolo del nuovo valore per la variabile
S68:        TEMP = (B(I) - SUM) / A(I, I)
            If Math.Abs(TEMP - X(I)) <= BIG Then GoTo S71
            BIG = Math.Abs(TEMP - X(I))
S71:        X(I) = TEMP
        Next I
        'Controllo della convergenza
        If BIG < ERR Then Return
        If ITER > Max Then GoTo S78
        ITER = ITER + 1
        GoTo S53
S78:    TextBox1.Text &= ("Numero di iterazioni superato") & vbCrLf
        Return
    End Sub

    Private Sub MATIN()
        ' Inizio del calcolo della matrice inversa
        For I = 1 To N
            R(I) = B(I)
        Next I
        A(1, 1) = 1 / A(1, 1)
        For M = 1 To NN
            K = M + 1
            For I = 1 To M
                B(I) = 0
                For J = 1 To M
                    B(I) = B(I) + A(I, J) * A(J, K)
                Next J
            Next I
            V = 0
            For I = 1 To M
                V = V + A(K, I) * B(I)
            Next I
            V = -V + A(K, K)
            A(K, K) = 1 / V
            For I = 1 To M
                A(I, K) = -B(I) * A(K, K)
            Next I
            For J = 1 To M
                C(J) = 0
                For I = 1 To M
                    C(J) = C(J) + A(K, I) * A(I, J)
                Next I
            Next J
            For J = 1 To M
                A(K, J) = -C(J) * A(K, K)
            Next J
            For I = 1 To M
                For J = 1 To M
                    A(I, J) = A(I, J) - B(I) * A(K, J)
                Next J
            Next I
        Next M
        ' Esegue le sostituzioni inverse
        D(1, 1) = 0
        For I = 1 To N
            For J = 1 To N
                D(I, J) = D(I, J) + A(I, J) * R(I)
            Next J
        Next I
        For J = 1 To N
            SUMX(J) = 0
            For I = 1 To N
                X(I) = D(I, J)
                SUMX(J) = X(I) + SUMX(J)
            Next I
        Next J
        For I = 1 To N
            J = I
            X(I) = SUMX(J)
        Next I
        Return
    End Sub
End Class
