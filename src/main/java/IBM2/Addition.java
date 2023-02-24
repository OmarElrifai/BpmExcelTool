package IBM2;

public class Addition {
    private float num1;
    private float num2;


    public void setNum2(float num2) {
        this.num2 = num2;
    }

    public void setNum1(float num1) {
        this.num1 = num1;
    }

    public float getNum1() {
        return this.num1;
    }

    public float getNum2() {
        return this.num2;
    }
    public float operation(float num1, float num2) {
        this.setNum1(num1);
        this.setNum2(num2);
        return this.getNum1() + this.getNum2();
    }

}
