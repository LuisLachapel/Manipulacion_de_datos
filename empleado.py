from sqlalchemy import Column, Float, Integer, String, create_engine
from sqlalchemy.orm import declarative_base, sessionmaker

engine = create_engine('sqlite:///Output/empleados_orm.db')

Base = declarative_base()

Session = sessionmaker(bind=engine)
session = Session()


class Empleado(Base):
   __tablename__ = 'empleados'
   id = Column(Integer, primary_key=True)
   nombre = Column(String)
   apellido = Column(String)
   edad = Column(Integer)
   salario = Column(Float)
   departamento = Column(String)

   def __repr__(self):
      return f"<Empleado(nombre='{self.nombre}', apellido='{self.apellido}', departamento='{self.departamento}')>"
   
"""
Base.metadata.create_all(engine)

empleado = Empleado(nombre = 'Carlos', apellido ='Mihazeru', edad= 25, salario = 50000, departamento ='TI' )
session.add(empleado)
session.commit()

empleados_ti = session.query(Empleado).filter_by(departamento='TI').all()
for empleado in empleados_ti:
    print(empleado)
"""